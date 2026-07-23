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
        // fp32: a variante q8 nao existe pra esse modelo (404) — fp32 e' o que funciona (~6s).
        // session_options: usa TODOS os nucleos + otimizacao de grafo (ajuda bastante em CPU fraca,
        // onde o onnxruntime-node as vezes usa poucas threads por padrao).
        const _cores = Math.max(2, (require('os').cpus() || []).length || 4);
        model = await mod.AutoModel.from_pretrained(MODEL_ID, {
            dtype: 'fp32',
            session_options: { graphOptimizationLevel: 'all', intraOpNumThreads: _cores, enableCpuMemArena: true }
        });
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

    // Output BiRefNet: [1, 1, H, W] — extrai raw data e faz sigmoid/scale/clamp em JS puro
    // (evita .sigmoid()/.to()/.squeeze() que quebram com "t.getValue is not a function")
    const maskTensor = out.output_image ?? out.logits ?? Object.values(out)[0];
    if (!maskTensor) throw new Error('Modelo nao retornou mascara valida');
    const raw = maskTensor.data || maskTensor.cpuData || maskTensor.buffer;
    if (!raw || !raw.length) throw new Error('Tensor sem dados extraiveis');
    const dims = maskTensor.dims || maskTensor.shape || [];
    const mW = dims[dims.length - 1];
    const mH = dims[dims.length - 2];
    if (!mW || !mH) throw new Error('Dims invalidas: ' + JSON.stringify(dims));
    const uint8 = new Uint8Array(mW * mH);
    for (let i = 0; i < mW * mH; i++) {
        const sig = 1 / (1 + Math.exp(-raw[i]));
        const v = sig * 255;
        uint8[i] = v < 0 ? 0 : v > 255 ? 255 : Math.round(v);
    }
    const rawMask = new RawImage(uint8, mW, mH, 1);
    const maskImage = await rawMask.resize(W, H);

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
