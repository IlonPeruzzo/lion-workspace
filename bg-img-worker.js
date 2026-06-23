// ═══════════════════════════════════════════════════════════════════
// BG Image Worker — roda BiRefNet em worker_thread isolado do main.
// Mantém o HTTP server (main process) responsivo durante inferência,
// evitando que o plugin Premiere marque "Desconectado" pelos pings
// que falham quando o main bloqueia >3s.
// ═══════════════════════════════════════════════════════════════════
'use strict';

const { parentPort } = require('worker_threads');
const fs = require('fs');
const path = require('path');

let mod = null;
let model = null;
let processor = null;

async function ensureModel() {
    if (mod && model && processor) return;
    if (!mod) {
        mod = await import('@huggingface/transformers');
    }
    const MODEL_ID = 'onnx-community/BiRefNet_lite-ONNX';
    if (!model) {
        model = await mod.AutoModel.from_pretrained(MODEL_ID, { dtype: 'fp32' });
    }
    if (!processor) {
        processor = await mod.AutoProcessor.from_pretrained(MODEL_ID);
    }
}

async function processImage(imagePath, outputPath) {
    await ensureModel();
    const sharp = require('sharp');
    const { RawImage } = mod;

    const meta = await sharp(imagePath).metadata();
    const W = meta.width;
    const H = meta.height;

    const img = await RawImage.read(imagePath);
    const inputs = await processor(img);
    const out = await model({ input_image: inputs.pixel_values });

    // Output BiRefNet: [1, 1, H, W] — squeeze(0) tira batch (RawImage espera 3D)
    const maskTensor = out.output_image ?? out.logits ?? Object.values(out)[0];
    const maskImage = await RawImage.fromTensor(
        maskTensor.sigmoid().mul(255).clamp(0, 255).to('uint8').squeeze(0)
    ).resize(W, H);

    // Compõe RGBA: RGB do original + alpha da máscara
    const origRgba = await sharp(imagePath).ensureAlpha().raw().toBuffer();
    const maskData = maskImage.data;
    const composed = Buffer.alloc(W * H * 4);
    for (let p = 0, j = 0; p < W * H * 4; p += 4, j++) {
        composed[p]     = origRgba[p];
        composed[p + 1] = origRgba[p + 1];
        composed[p + 2] = origRgba[p + 2];
        composed[p + 3] = maskData[j];
    }
    const pngBuf = await sharp(composed, { raw: { width: W, height: H, channels: 4 } })
        .png()
        .toBuffer();

    fs.writeFileSync(outputPath, pngBuf);
}

parentPort.on('message', async (msg) => {
    const { reqId, imagePath, outputPath } = msg;
    try {
        await processImage(imagePath, outputPath);
        parentPort.postMessage({ reqId, ok: true, outputPath });
    } catch (e) {
        parentPort.postMessage({ reqId, ok: false, error: (e && e.message) ? e.message : String(e) });
    }
});

// Sinaliza pro main que o worker iniciou (ainda não carregou o modelo)
parentPort.postMessage({ ready: true });
