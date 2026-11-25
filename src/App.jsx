import { useState, useRef } from "react";
import PptxGenJS from "pptxgenjs";
import heic2any from "heic2any";

const PRESETS = [1, 2, 4, 6, 9];
const W = 13.333, H = 7.5, MARGIN = 0.4, GAP = 0.2;

const isHeic = (f) => /heic|heif/i.test(f.type) || /\.(heic|heif)$/i.test(f.name);
const isVideo = (f) => f.type.startsWith("video/") || /\.(mp4|mov|m4v|webm|avi|mkv|3gp)$/i.test(f.name);

const loadImg = (src) =>
  new Promise((res, rej) => {
    const i = new Image();
    i.onload = () => res(i);
    i.onerror = rej;
    i.src = src;
  });

function toJpeg(source, w, h) {
  const s = Math.min(1, 1600 / Math.max(w, h));
  const c = document.createElement("canvas");
  c.width = Math.round(w * s);
  c.height = Math.round(h * s);
  const x = c.getContext("2d");
  x.fillStyle = "#fff";
  x.fillRect(0, 0, c.width, c.height);
  x.drawImage(source, 0, 0, c.width, c.height);
  return { data: c.toDataURL("image/jpeg", 0.85), w: c.width, h: c.height };
}

function frameFromVideo(url) {
  return new Promise((res, rej) => {
    const v = document.createElement("video");
    v.muted = true;
    v.playsInline = true;
    v.preload = "auto";
    v.onloadedmetadata = () => {
      v.currentTime = Math.min(1, (v.duration || 2) / 2);
    };
    v.onseeked = () => res(toJpeg(v, v.videoWidth, v.videoHeight));
    v.onerror = rej;
    v.src = url;
  });
}

async function processFile(file) {
  let url;
  try {
    if (isVideo(file)) {
      url = URL.createObjectURL(file);
      const f = await frameFromVideo(url);
      return { kind: "video", name: file.name, ...f };
    }
    let blob = file;
    if (isHeic(file)) {
      const out = await heic2any({ blob: file, toType: "image/jpeg", quality: 0.9 });
      blob = Array.isArray(out) ? out[0] : out;
    }
    url = URL.createObjectURL(blob);
    const img = await loadImg(url);
    const f = toJpeg(img, img.naturalWidth, img.naturalHeight);
    return { kind: "image", name: file.name, ...f };
  } catch (e) {
    console.warn("Skipped", file.name, e);
    return null;
  } finally {
    if (url) URL.revokeObjectURL(url);
  }
}

export default function App() {
  const [items, setItems] = useState([]);
  const [perSlide, setPerSlide] = useState(1);
  const [title, setTitle] = useState("My Presentation");
  const [busy, setBusy] = useState("");
  const [skipped, setSkipped] = useState(0);
  const [drag, setDrag] = useState(false);
  const inputRef = useRef();

  const n = Math.max(1, Math.min(25, Number(perSlide) || 1));
  const slideCount = Math.ceil(items.length / n);

  async function addFiles(fileList) {
    const files = Array.from(fileList);
    if (!files.length) return;
    let bad = 0;
    const added = [];
    for (let i = 0; i < files.length; i++) {
      setBusy(`Processing ${i + 1} of ${files.length}…`);
      const r = await processFile(files[i]);
      if (r) added.push({ id: crypto.randomUUID(), ...r });
      else bad++;
    }
    setItems((p) => [...p, ...added]);
    setSkipped(bad);
    setBusy("");
  }

  async function createPpt() {
    if (!items.length) return;
    setBusy("Creating your PPT…");
    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.title = title;
    const cols = Math.ceil(Math.sqrt(n));
    const rows = Math.ceil(n / cols);
    const cw = (W - 2 * MARGIN - (cols - 1) * GAP) / cols;
    const ch = (H - 2 * MARGIN - (rows - 1) * GAP) / rows;

    for (let s = 0; s < items.length; s += n) {
      const slide = pptx.addSlide();
      slide.background = { color: "FFFFFF" };
      items.slice(s, s + n).forEach((it, k) => {
        const r = Math.floor(k / cols), c = k % cols;
        const sc = Math.min(cw / it.w, ch / it.h);
        const iw = it.w * sc, ih = it.h * sc;
        const x = MARGIN + c * (cw + GAP) + (cw - iw) / 2;
        const y = MARGIN + r * (ch + GAP) + (ch - ih) / 2;
        slide.addImage({ data: it.data, x, y, w: iw, h: ih });
        if (it.kind === "video")
          slide.addText("▶ video", { x, y, w: 1, h: 0.35, fontSize: 11, color: "FFFFFF", fill: { color: "000000", transparency: 40 }, align: "center" });
      });
    }
    await pptx.writeFile({ fileName: `${title.trim() || "presentation"}.pptx` });
    setBusy("");
  }

  return (
    <div className="page">
      <header>
        <h1>📽️ Instant PPT Maker</h1>
        <p>Upload photos or videos, pick slides layout, get your PPT instantly.</p>
      </header>

      <main>
        <div
          className={`drop ${drag ? "over" : ""}`}
          onClick={() => inputRef.current.click()}
          onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={(e) => { e.preventDefault(); setDrag(false); addFiles(e.dataTransfer.files); }}
        >
          <input ref={inputRef} type="file" multiple hidden accept="image/*,video/*,.heic,.heif"
            onChange={(e) => { addFiles(e.target.files); e.target.value = ""; }} />
          <strong>Click or drag &amp; drop files here</strong>
          <span>JPG, PNG, JPEG, HEIC (iPhone), WEBP, GIF, MP4, MOV and more</span>
        </div>

        {busy && <p className="status">{busy}</p>}
        {skipped > 0 && <p className="warn">{skipped} file(s) could not be read and were skipped.</p>}

        <section className="card">
          <label className="field">
            Presentation name
            <input value={title} onChange={(e) => setTitle(e.target.value)} />
          </label>

          <div className="field">
            Photos / videos per slide
            <div className="chips">
              {PRESETS.map((p) => (
                <button key={p} className={n === p ? "chip on" : "chip"} onClick={() => setPerSlide(p)}>{p}</button>
              ))}
              <input className="custom" type="number" min="1" max="25" value={perSlide}
                onChange={(e) => setPerSlide(e.target.value)} title="Custom number" />
            </div>
          </div>

          <p className="summary">{items.length} file(s) → <b>{slideCount}</b> slide(s)</p>

          <button className="primary" disabled={!items.length || !!busy} onClick={createPpt}>
            ⬇ Create &amp; Download PPT
          </button>
          {items.length > 0 && <button className="link" onClick={() => setItems([])}>Clear all</button>}
        </section>

        {items.length > 0 && (
          <div className="grid">
            {items.map((it) => (
              <div className="thumb" key={it.id}>
                <img src={it.data} alt={it.name} />
                {it.kind === "video" && <span className="tag">▶</span>}
                <button onClick={() => setItems((p) => p.filter((x) => x.id !== it.id))} aria-label="Remove">×</button>
              </div>
            ))}
          </div>
        )}
      </main>

      <footer>
        Built with love ❤️ by Sriram H ·{" "}
        <a href="mailto:ramsri1702@gmail.com">ramsri1702@gmail.com</a>
      </footer>
    </div>
  );
}
