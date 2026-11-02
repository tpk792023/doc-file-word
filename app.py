"""Web app đọc nội dung file Word và xuất ra MP3 (giọng tiếng Việt)."""
import asyncio
import io
import re
import webbrowser
from datetime import datetime
from threading import Timer

import edge_tts
from docx import Document
from flask import Flask, jsonify, request, send_file, render_template_string
from gtts import gTTS

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB

VOICES = {
    "vi-VN-HoaiMyNeural": "Hoài My (nữ)",
    "vi-VN-NamMinhNeural": "Nam Minh (nam)",
    "google": "Google (nữ, dự phòng)",
}


def extract_text(file_stream):
    """Lấy chữ từ file .docx theo đúng thứ tự đoạn văn và bảng."""
    doc = Document(file_stream)
    parts = []
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            text = "".join(t.text or "" for t in child.iter() if t.tag.endswith("}t"))
            if text.strip():
                parts.append(text.strip())
        elif tag == "tbl":
            for row in child.iter():
                if not row.tag.endswith("}tr"):
                    continue
                cells = []
                for cell in row:
                    if cell.tag.endswith("}tc"):
                        ct = "".join(t.text or "" for t in cell.iter() if t.tag.endswith("}t"))
                        if ct.strip():
                            cells.append(ct.strip())
                if cells:
                    parts.append(", ".join(cells) + ".")
    return "\n\n".join(parts)


def split_text(text, limit=1500):
    """Chia văn bản dài thành nhiều đoạn nhỏ, cắt ở cuối đoạn văn hoặc cuối câu."""
    pieces = []
    for para in re.split(r"\n\s*\n", text):
        para = " ".join(para.split())
        while len(para) > limit:
            cut = max(para.rfind(s, 0, limit) for s in (". ", "! ", "? ", "; ", ", "))
            cut = cut + 1 if cut > 0 else limit
            pieces.append(para[:cut].strip())
            para = para[cut:].strip()
        if para:
            pieces.append(para)
    # Gộp các đoạn ngắn lại để đỡ phải gọi dịch vụ nhiều lần
    chunks, cur = [], ""
    for p in pieces:
        if cur and len(cur) + len(p) + 2 > limit:
            chunks.append(cur)
            cur = p
        else:
            cur = f"{cur}\n{p}" if cur else p
    if cur:
        chunks.append(cur)
    return chunks


async def edge_chunk(text, voice, rate, tries=2):
    """Giọng Microsoft đôi khi trả lỗi ngẫu nhiên nên thử lại vài lần."""
    for attempt in range(tries):
        data = bytearray()
        try:
            async with asyncio.timeout(15):
                async for chunk in edge_tts.Communicate(text, voice, rate=rate).stream():
                    if chunk["type"] == "audio":
                        data.extend(chunk["data"])
            if data:
                return bytes(data)
        except Exception:
            pass
        await asyncio.sleep(0.5)
    return None


def google_chunk(text, speed):
    """Giọng Google dự phòng (chỉ có giọng nữ, chỉnh được chậm/bình thường)."""
    buf = io.BytesIO()
    gTTS(text, lang="vi", slow=speed < 0).write_to_fp(buf)
    return buf.getvalue()


async def synthesize(text, voice, speed):
    """Trả về (file mp3, đã dùng giọng dự phòng hay chưa)."""
    rate = f"{speed:+d}%"
    parts = split_text(text)
    if voice != "google":
        buf = io.BytesIO()
        for part in parts:
            data = await edge_chunk(part, voice, rate)
            if data is None:
                break
            buf.write(data)
        else:
            buf.seek(0)
            return buf, False
    # Giọng Microsoft lỗi -> đọc lại toàn bộ bằng Google để cả file cùng một giọng
    buf = io.BytesIO()
    for part in parts:
        buf.write(google_chunk(part, speed))
    buf.seek(0)
    return buf, voice != "google"


@app.route("/")
def index():
    return render_template_string(PAGE, voices=VOICES)


@app.post("/api/extract")
def api_extract():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".docx"):
        return jsonify(error="Vui lòng chọn file Word đuôi .docx (file .doc cũ hãy mở bằng Word rồi Lưu thành .docx)."), 400
    try:
        text = extract_text(f.stream)
    except Exception as e:
        return jsonify(error=f"Không đọc được file: {e}"), 400
    if not text.strip():
        return jsonify(error="File không có chữ nào để đọc."), 400
    return jsonify(text=text, name=re.sub(r"\.docx$", "", f.filename, flags=re.I))


@app.post("/api/tts")
def api_tts():
    data = request.get_json(force=True)
    text = (data.get("text") or "").strip()
    voice = data.get("voice") if data.get("voice") in VOICES else "vi-VN-HoaiMyNeural"
    speed = int(data.get("speed") or 0)
    speed = max(-50, min(100, speed))
    name = re.sub(r'[\\/:*?"<>|]', "_", data.get("name") or "") or "doc-" + datetime.now().strftime("%Y%m%d-%H%M")
    if not text:
        return jsonify(error="Chưa có nội dung để đọc."), 400
    try:
        audio, fallback = asyncio.run(synthesize(text, voice, speed))
    except Exception as e:
        return jsonify(error=f"Không tạo được âm thanh (cần có Internet): {e}"), 500
    resp = send_file(audio, mimetype="audio/mpeg", as_attachment=True, download_name=f"{name}.mp3")
    resp.headers["X-Fallback"] = "1" if fallback else "0"
    return resp


PAGE = """<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Đọc File Word</title>
<style>
  :root { --bg:#f4f6fb; --card:#fff; --text:#1d2433; --muted:#667085; --accent:#2563eb; --accent2:#16a34a; --border:#d6dbe6; }
  * { box-sizing: border-box; }
  body { margin:0; font-family: "Segoe UI", Arial, sans-serif; background:var(--bg); color:var(--text); }
  .wrap { max-width: 900px; margin: 0 auto; padding: 24px 16px 48px; }
  h1 { margin: 0 0 4px; font-size: 26px; }
  .sub { color: var(--muted); margin: 0 0 20px; }
  .card { background:var(--card); border:1px solid var(--border); border-radius:14px; padding:20px; margin-bottom:16px; }
  .step { font-weight:600; margin-bottom:12px; }
  .drop { border:2px dashed var(--border); border-radius:12px; padding:28px; text-align:center; cursor:pointer; color:var(--muted); transition:.15s; }
  .drop.over, .drop:hover { border-color:var(--accent); background:#eef3ff; color:var(--accent); }
  .drop strong { color: var(--text); }
  textarea { width:100%; min-height:280px; border:1px solid var(--border); border-radius:10px; padding:12px; font-size:16px; line-height:1.6; font-family:inherit; resize:vertical; }
  .row { display:flex; flex-wrap:wrap; gap:16px; align-items:end; }
  label { display:block; font-size:14px; color:var(--muted); margin-bottom:6px; }
  select, input[type=text] { padding:10px; border:1px solid var(--border); border-radius:8px; font-size:15px; min-width:200px; }
  input[type=range] { width:200px; }
  .btns { display:flex; flex-wrap:wrap; gap:10px; margin-top:16px; }
  button { border:0; border-radius:10px; padding:12px 20px; font-size:16px; font-weight:600; cursor:pointer; color:#fff; background:var(--accent); }
  button.green { background:var(--accent2); }
  button.gray { background:#64748b; }
  button:disabled { opacity:.5; cursor:not-allowed; }
  audio { width:100%; margin-top:16px; }
  .msg { margin-top:12px; font-size:15px; }
  .msg.err { color:#dc2626; } .msg.ok { color:var(--accent2); }
  .count { color:var(--muted); font-size:13px; margin-top:6px; }
</style>
</head>
<body>
<div class="wrap">
  <h1>📄🔊 Đọc File Word</h1>
  <p class="sub">Tải file Word lên, nghe máy đọc nội dung và tải về file MP3.</p>

  <div class="card">
    <div class="step">Bước 1 — Chọn file Word (.docx)</div>
    <div class="drop" id="drop">
      <strong>Bấm vào đây</strong> hoặc kéo thả file Word vào ô này
      <div id="fname" style="margin-top:8px"></div>
    </div>
    <input type="file" id="file" accept=".docx" hidden>
  </div>

  <div class="card">
    <div class="step">Bước 2 — Kiểm tra nội dung (có thể sửa hoặc dán chữ trực tiếp vào đây)</div>
    <textarea id="text" placeholder="Nội dung file Word sẽ hiện ở đây..."></textarea>
    <div class="count" id="count">0 ký tự</div>
  </div>

  <div class="card">
    <div class="step">Bước 3 — Chọn giọng đọc và xuất MP3</div>
    <div class="row">
      <div><label>Giọng đọc</label>
        <select id="voice">{% for k, v in voices.items() %}<option value="{{k}}">{{v}}</option>{% endfor %}</select></div>
      <div><label>Tốc độ: <span id="speedLabel">Bình thường</span></label>
        <input type="range" id="speed" min="-50" max="100" step="10" value="0"></div>
      <div><label>Tên file MP3</label><input type="text" id="name" placeholder="ten-file"></div>
    </div>
    <div class="btns">
      <button id="play">▶ Nghe đọc</button>
      <button id="download" class="green">⬇ Tải file MP3</button>
      <button id="stop" class="gray">■ Dừng</button>
    </div>
    <div class="msg" id="msg"></div>
    <audio id="audio" controls hidden></audio>
  </div>
</div>
<script>
const $ = id => document.getElementById(id);
const drop = $("drop"), fileIn = $("file"), text = $("text"), audio = $("audio"), msg = $("msg");
let lastKey = null, lastBlob = null;
const OFFLINE = "Ứng dụng đã bị tắt. Hãy bấm đúp file CHAY-UNG-DUNG.bat để mở lại (giữ cửa sổ đen mở trong lúc dùng), rồi tải lại trang này.";

function say(t, cls) { msg.textContent = t; msg.className = "msg " + (cls || ""); }
function updateCount() { $("count").textContent = text.value.length.toLocaleString("vi-VN") + " ký tự"; }
text.addEventListener("input", updateCount);

$("speed").addEventListener("input", e => {
  const v = +e.target.value;
  $("speedLabel").textContent = v === 0 ? "Bình thường" : (v > 0 ? "Nhanh +" : "Chậm ") + v + "%";
});

drop.onclick = () => fileIn.click();
drop.ondragover = e => { e.preventDefault(); drop.classList.add("over"); };
drop.ondragleave = () => drop.classList.remove("over");
drop.ondrop = e => { e.preventDefault(); drop.classList.remove("over"); if (e.dataTransfer.files[0]) upload(e.dataTransfer.files[0]); };
fileIn.onchange = () => fileIn.files[0] && upload(fileIn.files[0]);

async function upload(f) {
  $("fname").textContent = "Đang đọc: " + f.name + "...";
  say("");
  const fd = new FormData(); fd.append("file", f);
  let r, j;
  try { r = await fetch("/api/extract", { method: "POST", body: fd }); j = await r.json(); }
  catch (e) { $("fname").textContent = ""; return say(OFFLINE, "err"); }
  if (!r.ok) { $("fname").textContent = ""; return say(j.error, "err"); }
  $("fname").textContent = "✔ " + f.name;
  text.value = j.text; $("name").value = j.name; updateCount();
  say("Đã lấy nội dung. Bấm “Nghe đọc” hoặc “Tải file MP3”.", "ok");
}

async function getAudio() {
  const body = { text: text.value, voice: $("voice").value, speed: +$("speed").value, name: $("name").value };
  const key = JSON.stringify([body.text, body.voice, body.speed]);
  if (key === lastKey && lastBlob) return lastBlob;
  if (!body.text.trim()) { say("Chưa có nội dung để đọc.", "err"); return null; }
  say("⏳ Đang tạo giọng đọc... (văn bản dài có thể mất vài chục giây)");
  $("play").disabled = $("download").disabled = true;
  try {
    const r = await fetch("/api/tts", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!r.ok) { const j = await r.json(); say(j.error, "err"); return null; }
    lastBlob = await r.blob(); lastKey = key;
    say(r.headers.get("X-Fallback") === "1"
      ? "✔ Xong! (Giọng Microsoft đang bận nên đã tự chuyển sang giọng Google dự phòng)"
      : "✔ Xong!", "ok");
    return lastBlob;
  } catch (e) { say(OFFLINE, "err"); return null; }
  finally { $("play").disabled = $("download").disabled = false; }
}

$("play").onclick = async () => {
  const b = await getAudio(); if (!b) return;
  audio.src = URL.createObjectURL(b); audio.hidden = false; audio.play();
};
$("download").onclick = async () => {
  const b = await getAudio(); if (!b) return;
  const a = document.createElement("a");
  a.href = URL.createObjectURL(b); a.download = ($("name").value.trim() || "doc-word") + ".mp3"; a.click();
};
$("stop").onclick = () => { audio.pause(); audio.currentTime = 0; };
</script>
</body>
</html>"""

if __name__ == "__main__":
    import os
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if os.environ.get("PORT"):  # Chạy trên máy chủ (Hugging Face)
        app.run(host="0.0.0.0", port=int(os.environ["PORT"]))
    else:  # Chạy trên máy tính của thầy
        Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5050")).start()
        print("Ứng dụng đang chạy tại http://127.0.0.1:5050  (đóng cửa sổ này để tắt)")
        app.run(host="127.0.0.1", port=5050)
