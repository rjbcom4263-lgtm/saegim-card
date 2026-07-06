// 새김 프록시 — NVIDIA NIM (이미지 생성 + 챗 완성)
// 배포: 웹 앱 / 실행: 나 / 액세스: 모든 사용자

function doPost(e) {
  let output;
  try {
    const params = JSON.parse(e.postData.contents);
    if (params.action === 'chat') {
      output = handleChat_(params);
    } else {
      output = handleNvidia_(params);
    }
  } catch (err) {
    output = { success: false, message: 'doPost 오류: ' + String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '새김 뱃지 프록시 동작 중' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleNvidia_(p) {
  const apiKey = (p.apiKey || '').trim();
  const model  = (p.model  || '').trim();
  const prompt = (p.prompt || '').trim();
  const seed   = parseInt(p.seed) || 0;

  if (!apiKey) return { success: false, message: 'NVIDIA API 키가 없습니다.' };
  if (!model)  return { success: false, message: 'model이 없습니다.' };
  if (!prompt) return { success: false, message: 'prompt가 없습니다.' };

  const isFlux = model.toLowerCase().includes('flux');
  let url, payload;

  if (isFlux) {
    // FLUX → integrate.api.nvidia.com (OpenAI 이미지 형식)
    url = 'https://integrate.api.nvidia.com/v1/images/generations';
    payload = { model: model, prompt: prompt, n: 1, size: '1024x1024', response_format: 'b64_json' };
    if (seed > 0) payload.seed = seed;
  } else {
    // SDXL / SD3 → ai.api.nvidia.com (Stability AI 형식)
    url = 'https://ai.api.nvidia.com/v1/genai/' + model;
    payload = {
      text_prompts: [{ text: prompt, weight: 1 }],
      cfg_scale: 7,
      sampler: 'DDIM',
      seed: seed > 0 ? seed : Math.floor(Math.random() * 99999),
      steps: 40,
      width: 1024,
      height: 1024
    };
  }

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code !== 200) {
    return { success: false, message: 'NVIDIA ' + code + ' [' + model + ']: ' + text.slice(0, 400) };
  }

  const data = JSON.parse(text);
  // FLUX → data[0].b64_json  /  SDXL → artifacts[0].base64
  const b64 = (data.data      && data.data[0])      ? data.data[0].b64_json
            : (data.artifacts && data.artifacts[0]) ? data.artifacts[0].base64
            : null;

  if (!b64) return { success: false, message: '이미지 없음: ' + text.slice(0, 200) };
  return { success: true, b64: b64 };
}

// ─── DeepSeek V4 Pro 챗 완성 ─────────────────────────────────────
function handleChat_(p) {
  var NVIDIA_KEY = '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';
  const apiKey   = NVIDIA_KEY !== 'YOUR_NVIDIA_KEY_HERE' ? NVIDIA_KEY : (p.apiKey || '').trim();
  const messages = p.messages  || [];
  const model    = p.model     || 'meta/llama-3.1-8b-instruct';

  if (!apiKey)          return { success: false, message: 'API 키가 없습니다.' };
  if (!messages.length) return { success: false, message: '메시지가 없습니다.' };

  const res = UrlFetchApp.fetch('https://integrate.api.nvidia.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      model: model,
      messages: messages,
      temperature: 0.7,
      max_tokens: 512,
      stream: false
    }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code !== 200) {
    return { success: false, message: 'API ' + code + ': ' + text.slice(0, 400) };
  }

  const data    = JSON.parse(text);
  const content = data.choices && data.choices[0] ? data.choices[0].message.content : null;
  if (!content) return { success: false, message: '응답 없음: ' + text.slice(0, 200) };
  return { success: true, content: content };
}
