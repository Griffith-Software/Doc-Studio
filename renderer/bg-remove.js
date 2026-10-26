'use strict';
// Remove background: the U²-Net-p model (renderer/vendor/bgremove, Apache-2.0) run by ONNX Runtime's WebAssembly
// build (renderer/vendor/ort, MIT), entirely on this PC. mask(picture) returns a 320×320 canvas whose alpha is how
// sure the model is that each spot belongs to the main subject; the image editor scales it to the picture.

const BgRemove = (() => {
  const SIZE = 320;
  let ready = null;

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error("The background remover couldn't be loaded."));
      document.head.append(s);
    });
  }
  // Loaded on first use (about 15 MB), then kept.
  function load() {
    if (!ready) {
      ready = (async () => {
        if (typeof ort === 'undefined') await loadScript('vendor/ort/ort.wasm.min.js');
        ort.env.wasm.numThreads = 1; // threads need cross-origin isolation, which a local page doesn't have
        ort.env.wasm.wasmPaths = { mjs: new URL('vendor/ort/ort-wasm-simd-threaded.mjs', location.href).href };
        ort.env.wasm.wasmBinary = await desk.asset('vendor/ort/ort-wasm-simd-threaded.wasm');
        const model = await desk.asset('vendor/bgremove/u2netp.onnx');
        return ort.InferenceSession.create(new Uint8Array(model), { executionProviders: ['wasm'], graphOptimizationLevel: 'all' });
      })();
      ready.catch(() => { ready = null; });
    }
    return ready;
  }

  async function mask(picture) {
    const session = await load();
    const c = document.createElement('canvas');
    c.width = SIZE; c.height = SIZE;
    const g = c.getContext('2d', { willReadFrequently: true });
    g.fillStyle = '#fff';
    g.fillRect(0, 0, SIZE, SIZE);
    g.drawImage(picture, 0, 0, SIZE, SIZE);
    const px = g.getImageData(0, 0, SIZE, SIZE).data;
    // The model's input: RGB planes, scaled by the brightest value, then normalised (ImageNet mean and spread).
    let max = 1;
    for (let i = 0; i < px.length; i += 4) max = Math.max(max, px[i], px[i + 1], px[i + 2]);
    const mean = [0.485, 0.456, 0.406], std = [0.229, 0.224, 0.225], n = SIZE * SIZE;
    const input = new Float32Array(3 * n);
    for (let p = 0, i = 0; p < n; p++, i += 4) for (let k = 0; k < 3; k++) input[k * n + p] = (px[i + k] / max - mean[k]) / std[k];
    const out = await session.run({ [session.inputNames[0]]: new ort.Tensor('float32', input, [1, 3, SIZE, SIZE]) });
    const pred = out[session.outputNames[0]].data;
    let lo = Infinity, hi = -Infinity;
    for (let p = 0; p < n; p++) { if (pred[p] < lo) lo = pred[p]; if (pred[p] > hi) hi = pred[p]; }
    const m = document.createElement('canvas');
    m.width = SIZE; m.height = SIZE;
    const mg = m.getContext('2d');
    const id = mg.createImageData(SIZE, SIZE);
    for (let p = 0; p < n; p++) id.data[p * 4 + 3] = Math.round((pred[p] - lo) / (hi - lo || 1) * 255);
    mg.putImageData(id, 0, 0);
    return m;
  }

  return { load, mask };
})();
