// xlsx（zip）を読んで、一部のファイルだけ書き換えて、新しい zip を作る。
// 外のライブラリは使わない：展開はブラウザの DecompressionStream（deflate-raw）。
// 書き換えないファイルは圧縮されたまま写し、書き換えたファイルだけ無圧縮（STORE）で入れる（Excel はどちらも読める）。
(function () {
  const td = new TextDecoder(), te = new TextEncoder();

  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(u8) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  async function inflate(u8) {
    const s = new Blob([u8]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(s).arrayBuffer());
  }

  // 中央ディレクトリを読む → [{name, method, crc, csize, usize, data(圧縮されたまま), extra}]
  function readZip(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(buf);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("xlsx（zip）の形ではありません");
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const out = [];
    for (let k = 0; k < count; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error("zip の目次が壊れています");
      const flags = dv.getUint16(p + 8, true), method = dv.getUint16(p + 10, true);
      const time = dv.getUint16(p + 12, true), date = dv.getUint16(p + 14, true);
      const crc = dv.getUint32(p + 16, true), csize = dv.getUint32(p + 20, true), usize = dv.getUint32(p + 24, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = td.decode(u8.subarray(p + 46, p + 46 + nlen));
      const lnlen = dv.getUint16(lho + 26, true), lxlen = dv.getUint16(lho + 28, true);
      const start = lho + 30 + lnlen + lxlen;
      out.push({ name, flags, method, time, date, crc, csize, usize, data: u8.subarray(start, start + csize) });
      p += 46 + nlen + xlen + clen;
    }
    return out;
  }

  async function text(entry) {
    const raw = entry.method === 0 ? entry.data : await inflate(entry.data);
    return td.decode(raw);
  }

  function setText(entry, s) {
    const u8 = te.encode(s);
    entry.method = 0; entry.flags = 0x0800; entry.data = u8;
    entry.crc = crc32(u8); entry.csize = entry.usize = u8.length;
  }

  function writeZip(entries) {
    const parts = [], central = [];
    let off = 0;
    for (const e of entries) {
      const name = te.encode(e.name);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, e.flags & 0x0800, true);
      h.setUint16(8, e.method, true); h.setUint16(10, e.time, true); h.setUint16(12, e.date, true);
      h.setUint32(14, e.crc, true); h.setUint32(18, e.csize, true); h.setUint32(22, e.usize, true);
      h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(new Uint8Array(h.buffer), name, e.data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true);
      c.setUint16(8, e.flags & 0x0800, true); c.setUint16(10, e.method, true);
      c.setUint16(12, e.time, true); c.setUint16(14, e.date, true);
      c.setUint32(16, e.crc, true); c.setUint32(20, e.csize, true); c.setUint32(24, e.usize, true);
      c.setUint16(28, name.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), name);
      off += 30 + name.length + e.data.length;
    }
    const csize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
    end.setUint32(12, csize, true); end.setUint32(16, off, true);
    return new Blob([...parts, ...central, new Uint8Array(end.buffer)],
      { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  // ── シートの XML のセルを書き換える ──
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const colNum = ref => ref.replace(/\d+/g, "").split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

  // value：文字（inlineStr）・数（v）・null（中を空にする。書式は残す）
  function setCell(xml, ref, value) {
    const re = new RegExp(`<c r="${ref}"((?:\\s+(?!r=)[a-zA-Z:]+="[^"]*")*)\\s*(?:/>|>[\\s\\S]*?</c>)`);
    const m = xml.match(re);
    let attrs = "";
    if (m) attrs = m[1].replace(/\s+t="[^"]*"/, "");
    let cell;
    if (value === null || value === undefined || value === "") cell = `<c r="${ref}"${attrs}/>`;
    else if (typeof value === "number") cell = `<c r="${ref}"${attrs}><v>${value}</v></c>`;
    else cell = `<c r="${ref}"${attrs} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
    if (m) return xml.replace(re, cell);
    // セルが無い：その行の中の、列の順の位置に入れる
    const row = ref.replace(/\D+/g, "");
    const rre = new RegExp(`<row r="${row}"[^>]*?(?:/>|>([\\s\\S]*?)</row>)`);
    const rm = xml.match(rre);
    if (!rm) throw new Error(`行 ${row} がありません（${ref}）`);
    if (rm[0].endsWith("/>")) return xml.replace(rre, rm[0].replace(/\/>$/, `>${cell}</row>`));
    const col = colNum(ref);
    const inner = rm[1].replace(/<c r="([A-Z]+)\d+"/g, (s0) => s0);
    let pos = inner.length;
    const cre = /<c r="([A-Z]+)\d+"/g;
    let cm;
    while ((cm = cre.exec(inner))) { if (colNum(cm[1]) > col) { pos = cm.index; break; } }
    const newInner = inner.slice(0, pos) + cell + inner.slice(pos);
    return xml.replace(rre, rm[0].replace(rm[1], newInner));
  }

  window.XlsxEdit = { readZip, text, setText, writeZip, setCell, crc32 };
})();
