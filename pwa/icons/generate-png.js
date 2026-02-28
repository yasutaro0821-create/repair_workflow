/**
 * Node.js script to generate PWA icons as PNG
 * Requires no external dependencies - creates valid PNG files
 */
const fs = require('fs');
const { createCanvas } = (() => {
  // Simple canvas-like interface using raw pixel manipulation
  return {
    createCanvas: (w, h) => {
      const pixels = Buffer.alloc(w * h * 4); // RGBA
      return {
        width: w,
        height: h,
        pixels,
        setPixel(x, y, r, g, b, a = 255) {
          if (x < 0 || x >= w || y < 0 || y >= h) return;
          const i = (y * w + x) * 4;
          pixels[i] = r;
          pixels[i + 1] = g;
          pixels[i + 2] = b;
          pixels[i + 3] = a;
        },
        fillRect(x0, y0, rw, rh, r, g, b) {
          for (let y = y0; y < y0 + rh && y < h; y++) {
            for (let x = x0; x < x0 + rw && x < w; x++) {
              this.setPixel(x, y, r, g, b);
            }
          }
        },
        fillCircle(cx, cy, radius, r, g, b) {
          for (let y = cy - radius; y <= cy + radius; y++) {
            for (let x = cx - radius; x <= cx + radius; x++) {
              if ((x - cx) ** 2 + (y - cy) ** 2 <= radius ** 2) {
                this.setPixel(Math.round(x), Math.round(y), r, g, b);
              }
            }
          }
        },
        fillRoundedRect(x0, y0, rw, rh, radius, r, g, b) {
          // Fill main body
          this.fillRect(x0 + radius, y0, rw - 2 * radius, rh, r, g, b);
          this.fillRect(x0, y0 + radius, rw, rh - 2 * radius, r, g, b);
          // Fill corners
          this.fillCircle(x0 + radius, y0 + radius, radius, r, g, b);
          this.fillCircle(x0 + rw - radius, y0 + radius, radius, r, g, b);
          this.fillCircle(x0 + radius, y0 + rh - radius, radius, r, g, b);
          this.fillCircle(x0 + rw - radius, y0 + rh - radius, radius, r, g, b);
        },
        drawLine(x0, y0, x1, y1, thickness, r, g, b) {
          const dx = x1 - x0, dy = y1 - y0;
          const len = Math.sqrt(dx * dx + dy * dy);
          const steps = Math.ceil(len * 2);
          for (let i = 0; i <= steps; i++) {
            const t = i / steps;
            const cx = x0 + dx * t;
            const cy = y0 + dy * t;
            this.fillCircle(cx, cy, thickness / 2, r, g, b);
          }
        },
        toPNG() {
          return encodePNG(this.pixels, w, h);
        }
      };
    }
  };
})();

// Minimal PNG encoder
function encodePNG(pixels, width, height) {
  const crc32Table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    crc32Table[i] = c;
  }
  function crc32(buf) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < buf.length; i++) c = crc32Table[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // Create raw image data (filter type 0 = None for each row)
  const rawData = [];
  for (let y = 0; y < height; y++) {
    rawData.push(0); // filter byte
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 4;
      rawData.push(pixels[i], pixels[i + 1], pixels[i + 2], pixels[i + 3]);
    }
  }

  // Deflate using zlib
  const zlib = require('zlib');
  const compressed = zlib.deflateSync(Buffer.from(rawData));

  function chunk(type, data) {
    const typeBuf = Buffer.from(type, 'ascii');
    const lenBuf = Buffer.alloc(4);
    lenBuf.writeUInt32BE(data.length);
    const crcData = Buffer.concat([typeBuf, data]);
    const crcBuf = Buffer.alloc(4);
    crcBuf.writeUInt32BE(crc32(crcData));
    return Buffer.concat([lenBuf, typeBuf, data, crcBuf]);
  }

  // IHDR chunk
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type: RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace

  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), // PNG signature
    chunk('IHDR', ihdr),
    chunk('IDAT', compressed),
    chunk('IEND', Buffer.alloc(0))
  ]);
}

function generateIcon(size) {
  const canvas = createCanvas(size, size);
  const s = size / 192; // scale factor

  // Blue background with rounded corners
  const radius = Math.round(32 * s);
  canvas.fillRoundedRect(0, 0, size, size, radius, 26, 115, 232);

  // Draw wrench icon (white)
  const cx = size / 2;
  const cy = size / 2;
  const t = Math.round(8 * s); // line thickness

  // Wrench handle (diagonal)
  canvas.drawLine(
    cx - 30 * s, cy + 30 * s,
    cx + 20 * s, cy - 20 * s,
    t, 255, 255, 255
  );

  // Wrench head (open jaw)
  canvas.drawLine(
    cx + 20 * s, cy - 20 * s,
    cx + 35 * s, cy - 40 * s,
    t, 255, 255, 255
  );
  canvas.drawLine(
    cx + 20 * s, cy - 20 * s,
    cx + 45 * s, cy - 25 * s,
    t, 255, 255, 255
  );

  // Wrench bottom ring
  canvas.fillCircle(
    cx - 30 * s, cy + 30 * s,
    12 * s, 255, 255, 255
  );
  canvas.fillCircle(
    cx - 30 * s, cy + 30 * s,
    6 * s, 26, 115, 232
  );

  // Small building accent at bottom
  const bx = cx + 15 * s;
  const by = cy + 50 * s;
  canvas.fillRect(
    Math.round(bx - 20 * s), Math.round(by - 15 * s),
    Math.round(40 * s), Math.round(20 * s),
    255, 255, 255
  );
  // Roof
  for (let i = 0; i < Math.round(8 * s); i++) {
    canvas.fillRect(
      Math.round(bx - 20 * s - i), Math.round(by - 15 * s - i),
      Math.round(40 * s + 2 * i), 1,
      255, 255, 255
    );
  }
  // Windows
  const ww = Math.round(6 * s);
  const wh = Math.round(6 * s);
  for (let wx = -1; wx <= 1; wx++) {
    canvas.fillRect(
      Math.round(bx + wx * 12 * s - ww / 2),
      Math.round(by - 8 * s),
      ww, wh,
      26, 115, 232
    );
  }

  return canvas.toPNG();
}

// Generate both sizes
const icon192 = generateIcon(192);
const icon512 = generateIcon(512);

const dir = __dirname;
fs.writeFileSync(`${dir}/icon-192.png`, icon192);
fs.writeFileSync(`${dir}/icon-512.png`, icon512);

console.log(`icon-192.png: ${icon192.length} bytes`);
console.log(`icon-512.png: ${icon512.length} bytes`);
console.log('Icons generated successfully!');
