const video = document.getElementById("camera");
const canvas = document.getElementById("overlay");
const ctx = canvas.getContext("2d");

const statusEl = document.getElementById("status");
const movementNameEl = document.getElementById("movementName");
const sideNameEl = document.getElementById("sideName");
const sideRowEl = document.getElementById("sideRow");
const rotaryRowEl = document.getElementById("rotaryRow");
const rotaryPhaseEl = document.getElementById("rotaryPhase");
const scoreTextEl = document.getElementById("scoreText");
const reasonTextEl = document.getElementById("reasonText");
const recordBtn = document.getElementById("recordBtn");
const painBtn = document.getElementById("painBtn");
const sideButtons = document.getElementById("sideButtons");
const movementList = document.getElementById("movementList");
const completedText = document.getElementById("completedText");
const samplesText = document.getElementById("samplesText");
const fpsText = document.getElementById("fpsText");
const folderNameText = document.getElementById("folderNameText");
const changeFolderBtn = document.getElementById("changeFolderBtn");
const folderDialog = document.getElementById("folderDialog");
const folderNameInput = document.getElementById("folderNameInput");
const folderStartBtn = document.getElementById("folderStartBtn");
const saveDialog = document.getElementById("saveDialog");
const testNameInput = document.getElementById("testNameInput");
const saveBtn = document.getElementById("saveBtn");
const downloadDialog = document.getElementById("downloadDialog");
const downloadLink = document.getElementById("downloadLink");
const downloadInfo = document.getElementById("downloadInfo");
const cloudUploadBtn = document.getElementById("cloudUploadBtn");
const finishBtn = document.getElementById("finishBtn");
const switchCameraBtn = document.getElementById("switchCameraBtn");
const zoomSlider = document.getElementById("zoomSlider");
const zoomText = document.getElementById("zoomText");
const GOOGLE_UPLOAD_URL = "https://script.google.com/macros/s/AKfycbxQ9cKAp--d4CV_Vf0NxVXVaz2t8MBpxKSPaBDtvTmsH14HIaWeCTIrEBaJ8ScIT5cEsg/exec";

const movements = [
  { id: "deep_squat", name: "深蹲", bilateral: false },
  { id: "hurdle_step", name: "跨欄步", bilateral: true },
  { id: "inline_lunge", name: "直線弓箭步", bilateral: true },
  { id: "shoulder_mobility", name: "肩關節活動度", bilateral: true },
  { id: "active_straight_leg_raise", name: "主動直膝抬腿", bilateral: true },
  { id: "trunk_stability_pushup", name: "軀幹穩定伏地挺身", bilateral: false },
  { id: "rotary_stability", name: "旋轉穩定性", bilateral: true },
];

const sideNames = { left: "左側", right: "右側" };
let pose;
let cameraReady = false;
let recording = false;
let rows = [];
let recordStart = 0;
let sampleCount = 0;
let measuredFps = 0;
let lastFrameTime = performance.now();
let painReported = false;
let manualOverrideScore = null;
let movementIndex = 0;
let currentSide = "left";
let rotaryPhase = "same";
let sessionResults = {};
let sessionSideResults = {};
let lastPoseLandmarks = null;
let processing = false;
let cameraFacingMode = "environment";
let cameraStream = null;
let pendingDownloadUrl = null;
let pendingExcelBlob = null;
let pendingExcelFileName = "";
let desiredZoom = 1;
let cloudFolderName = localStorage.getItem("fmsCloudFolderName") || "";

function currentMovement() {
  return movements[movementIndex];
}

function setStatus(text) {
  statusEl.textContent = text;
}

function updateFolderUi() {
  folderNameText.textContent = cloudFolderName || "未設定";
}

async function createCloudFolder(folderName) {
  const formData = new FormData();
  formData.append("action", "createFolder");
  formData.append("folderName", folderName);
  await fetch(GOOGLE_UPLOAD_URL, {
    method: "POST",
    body: formData,
    mode: "no-cors",
  });
}

async function enterCloudFolder(folderName) {
  const nextName = sanitizeName(folderName);
  cloudFolderName = nextName;
  localStorage.setItem("fmsCloudFolderName", nextName);
  updateFolderUi();
  folderDialog.close();
  setStatus(`目前資料夾：${nextName}`);
  try {
    await createCloudFolder(nextName);
  } catch (error) {
    console.warn("Cloud folder creation request failed.", error);
    setStatus(`已選擇資料夾：${nextName}`);
  }
}

function updateZoomUi(value, disabled = false, label = "") {
  zoomSlider.disabled = disabled;
  zoomSlider.value = String(value);
  zoomText.textContent = label || `${Number(value).toFixed(1)}x`;
}

async function applyCameraZoom(value) {
  const track = cameraStream?.getVideoTracks()[0];
  if (!track) return;
  const capabilities = track.getCapabilities?.() || {};
  if (!("zoom" in capabilities)) {
    updateZoomUi(1, true, "不支援");
    return;
  }
  const min = capabilities.zoom.min ?? 1;
  const max = capabilities.zoom.max ?? 1;
  const zoom = Math.min(max, Math.max(min, Number(value)));
  desiredZoom = zoom;
  await track.applyConstraints({ advanced: [{ zoom }] });
  updateZoomUi(zoom, false);
}

async function configureCameraZoom() {
  const track = cameraStream?.getVideoTracks()[0];
  const capabilities = track?.getCapabilities?.() || {};
  if (!track || !("zoom" in capabilities)) {
    desiredZoom = 1;
    updateZoomUi(1, true, "不支援");
    return;
  }
  const min = capabilities.zoom.min ?? 1;
  const max = capabilities.zoom.max ?? 1;
  const step = capabilities.zoom.step || 0.1;
  zoomSlider.min = String(min);
  zoomSlider.max = String(max);
  zoomSlider.step = String(step);
  desiredZoom = Math.min(max, Math.max(min, desiredZoom || min));
  await applyCameraZoom(desiredZoom);
}

function dist(a, b) {
  return Math.hypot(a.x - b.x, a.y - b.y);
}

function pointLineDistance(point, a, b) {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  if (dx === 0 && dy === 0) return dist(point, a);
  return Math.abs(dy * point.x - dx * point.y + b.x * a.y - b.y * a.x) / Math.hypot(dx, dy);
}

function average(values) {
  const valid = values.filter((value) => Number.isFinite(value));
  return valid.length ? valid.reduce((sum, value) => sum + value, 0) / valid.length : "";
}

function angle3d(a, b, c) {
  const ba = { x: a.x - b.x, y: a.y - b.y, z: a.z - b.z };
  const bc = { x: c.x - b.x, y: c.y - b.y, z: c.z - b.z };
  const dot = ba.x * bc.x + ba.y * bc.y + ba.z * bc.z;
  const magA = Math.hypot(ba.x, ba.y, ba.z);
  const magB = Math.hypot(bc.x, bc.y, bc.z);
  if (!magA || !magB) return "";
  const value = Math.max(-1, Math.min(1, dot / (magA * magB)));
  return Math.round((Math.acos(value) * 180 / Math.PI) * 1000) / 1000;
}

function visible(lm, indices, threshold = 0.5) {
  return indices.every((index) => (lm[index].visibility ?? 1) >= threshold);
}

function midpoint(lm, a, b) {
  return { x: (lm[a].x + lm[b].x) / 2, y: (lm[a].y + lm[b].y) / 2 };
}

function shoulderWidth(lm) {
  return Math.abs(lm[11].x - lm[12].x) || 1;
}

function normalized(value, reference) {
  return Math.abs(value) / (reference || 1);
}

function torsoLean(lm) {
  const shoulders = midpoint(lm, 11, 12);
  const hips = midpoint(lm, 23, 24);
  return Math.atan2(Math.abs(shoulders.x - hips.x), Math.abs(shoulders.y - hips.y)) * 180 / Math.PI;
}

function poseAngles(lm) {
  return {
    left_elbow: angle3d(lm[11], lm[13], lm[15]),
    right_elbow: angle3d(lm[12], lm[14], lm[16]),
    left_hip: angle3d(lm[11], lm[23], lm[25]),
    right_hip: angle3d(lm[12], lm[24], lm[26]),
    left_knee: angle3d(lm[23], lm[25], lm[27]),
    right_knee: angle3d(lm[24], lm[26], lm[28]),
  };
}

function result(score, reason) {
  return { score, reason };
}

function issueText(label, value, good, ok) {
  if (value === "") return "";
  if (value <= good) return "";
  if (value <= ok) return `${label}輕微偏移`;
  return `${label}明顯偏移`;
}

function frontalMetrics(lm, side) {
  const [hip, knee, ankle, heel, foot] = side === "left" ? [23, 25, 27, 29, 31] : [24, 26, 28, 30, 32];
  const width = shoulderWidth(lm);
  const shoulders = midpoint(lm, 11, 12);
  const hips = midpoint(lm, 23, 24);
  const footTurn = visible(lm, [heel, foot]) ? normalized(lm[foot].x - lm[heel].x, width) : "";
  return {
    hip_shift: normalized(hips.x - shoulders.x, width),
    pelvis_level: Math.abs(lm[23].y - lm[24].y),
    knee_line: pointLineDistance(lm[knee], lm[hip], lm[ankle]),
    ankle_shift: normalized(lm[ankle].x - lm[hip].x, width),
    foot_turn: footTurn,
  };
}

function evaluateDeepSquat(lm, pain) {
  if (pain) return result(0, "動作過程有疼痛，依 FMS 規則給 0 分");
  if (!visible(lm, [11, 12, 15, 16, 23, 24, 25, 26, 27, 28])) return result(1, "身體關節點辨識不清楚，無法完整判斷深蹲動作");
  const angles = poseAngles(lm);
  const avgKnee = average([angles.left_knee, angles.right_knee]);
  const shoulders = midpoint(lm, 11, 12);
  const hips = midpoint(lm, 23, 24);
  const hipShift = normalized(hips.x - shoulders.x, shoulderWidth(lm));
  const depth = ((lm[23].y + lm[24].y) / 2) - ((lm[25].y + lm[26].y) / 2);
  const leftTrack = pointLineDistance(lm[25], lm[23], lm[27]);
  const rightTrack = pointLineDistance(lm[26], lm[24], lm[28]);
  const kneeTrack = (leftTrack + rightTrack) / 2;
  const asym = Math.abs((angles.left_knee || 0) - (angles.right_knee || 0));
  const armsOverhead = lm[15].y < lm[11].y && lm[16].y < lm[12].y;
  const deep = depth >= 0.02 || (Number.isFinite(avgKnee) && avgKnee <= 110);
  const moderate = depth >= -0.02 || (Number.isFinite(avgKnee) && avgKnee <= 135);
  if (deep && kneeTrack <= 0.06 && hipShift <= 0.12 && asym <= 12 && armsOverhead) {
    return result(3, "深蹲深度足夠，髖中心穩定，膝蓋軌跡與左右對稱良好，雙手維持在肩膀上方");
  }
  if (moderate && kneeTrack <= 0.11 && hipShift <= 0.22 && asym <= 22) {
    return result(2, "可以完成深蹲，但有代償現象，例如深度、髖中心側移、膝蓋軌跡或左右對稱未達最佳");
  }
  return result(1, "無法完成符合標準的深蹲動作，可能深度不足、髖中心側移、膝蓋偏移或左右不對稱明顯");
}

function evaluateHurdleStep(lm, pain, side) {
  if (pain) return result(0, `${sideNames[side]}動作過程有疼痛，依 FMS 規則給 0 分`);
  const [hip, knee, ankle, heel, foot] = side === "left" ? [23, 25, 27, 29, 31] : [24, 26, 28, 30, 32];
  if (!visible(lm, [11, 12, 23, 24, hip, knee, ankle])) return result(1, `${sideNames[side]}關節點辨識不清楚，無法完整判斷跨欄步`);
  const kneeLift = ((lm[23].y + lm[24].y) / 2) - lm[knee].y;
  const metrics = frontalMetrics(lm, side);
  const footVisible = visible(lm, [heel, foot]);
  const issues = [];
  if (kneeLift < 0.04) issues.push("抬腿高度不足");
  ["骨盆高度", "髖中心左右位移", "髖膝踝排列", "腳踝內外偏", "足尖方向"].forEach((label, idx) => {
    const values = [metrics.pelvis_level, metrics.hip_shift, metrics.knee_line, metrics.ankle_shift, metrics.foot_turn];
    const good = [0.035, 0.12, 0.045, 0.16, 0.16][idx];
    const ok = [0.07, 0.22, 0.09, 0.25, 0.24][idx];
    const text = issueText(label, values[idx], good, ok);
    if (text) issues.push(text);
  });
  const footOk = metrics.foot_turn === "" || metrics.foot_turn <= 0.16;
  const footModerate = metrics.foot_turn === "" || metrics.foot_turn <= 0.24;
  const good = kneeLift >= 0.04 && metrics.pelvis_level <= 0.035 && metrics.hip_shift <= 0.12 && metrics.knee_line <= 0.045 && metrics.ankle_shift <= 0.16 && footOk;
  const ok = kneeLift >= 0 && metrics.pelvis_level <= 0.07 && metrics.hip_shift <= 0.22 && metrics.knee_line <= 0.09 && metrics.ankle_shift <= 0.25 && footModerate;
  if (good) return result(3, `${sideNames[side]}跨步高度足夠，骨盆水平，髖膝踝排列與腳踝控制良好`);
  if (ok) return result(2, `${sideNames[side]}可以完成跨欄步，但有代償：${issues.join("、") || "輕微代償"}`);
  return result(1, `${sideNames[side]}跨欄步不穩定：${issues.join("、") || "下肢或骨盆控制不穩"}${footVisible ? "" : "，且足部關節點不夠清楚"}`);
}

function evaluateInlineLunge(lm, pain, side) {
  if (pain) return result(0, `${sideNames[side]}動作過程有疼痛，依 FMS 規則給 0 分`);
  if (!visible(lm, [11, 12, 23, 24, 25, 26, 27, 28])) return result(1, `${sideNames[side]}關節點辨識不清楚，無法完整判斷直線弓箭步`);
  const angles = poseAngles(lm);
  const knee = side === "left" ? angles.left_knee : angles.right_knee;
  const hip = side === "left" ? angles.left_hip : angles.right_hip;
  const pelvis = Math.abs(lm[23].y - lm[24].y);
  const torso = torsoLean(lm);
  if (knee <= 125 && torso <= 20 && pelvis <= 0.05) return result(3, `${sideNames[side]}弓箭步深度與軀幹控制良好，下肢穩定`);
  if (knee <= 145 && torso <= 35 && hip <= 150) return result(2, `${sideNames[side]}可以完成弓箭步，但深度、軀幹穩定或下肢控制有代償`);
  return result(1, `${sideNames[side]}無法完成符合標準的直線弓箭步，可能深度不足或穩定度不足`);
}

function evaluateShoulderMobility(lm, pain, side) {
  if (pain) return result(0, `${sideNames[side]}肩關節活動時有疼痛，依 FMS 規則給 0 分`);
  const [wrist, shoulder, hip, otherWrist] = side === "left" ? [15, 11, 23, 16] : [16, 12, 24, 15];
  if (!visible(lm, [11, 12, wrist, otherWrist, hip])) return result(1, `${sideNames[side]}肩、手腕或軀幹關節點辨識不清楚`);
  const width = shoulderWidth(lm);
  const wristDistance = dist(lm[wrist], lm[otherWrist]) / width;
  const reached = lm[wrist].y < lm[shoulder].y || lm[wrist].y > lm[hip].y;
  if (wristDistance <= 1.0 && reached) return result(3, `${sideNames[side]}肩關節活動度良好，雙手距離接近`);
  if (wristDistance <= 1.8 && reached) return result(2, `${sideNames[side]}可以完成肩關節活動度動作，但雙手距離較大`);
  return result(1, `${sideNames[side]}肩關節活動度不足，雙手距離過大或無法到達測試位置`);
}

function evaluateLegRaise(lm, pain, side) {
  if (pain) return result(0, `${sideNames[side]}抬腿時有疼痛，依 FMS 規則給 0 分`);
  const ankle = side === "left" ? 27 : 28;
  if (!visible(lm, [23, 24, 25, 26, ankle])) return result(1, `${sideNames[side]}髖、膝或踝關節點辨識不清楚`);
  const kneeAngle = side === "left" ? poseAngles(lm).left_knee : poseAngles(lm).right_knee;
  const kneeStraight = kneeAngle >= 155;
  const hipY = (lm[23].y + lm[24].y) / 2;
  const kneeY = (lm[25].y + lm[26].y) / 2;
  const raiseHeight = kneeY - lm[ankle].y;
  if (kneeStraight && raiseHeight >= 0.08) return result(3, `${sideNames[side]}抬腿高度足夠且膝蓋維持伸直`);
  if (raiseHeight >= -0.02 || hipY - lm[ankle].y > 0) return result(2, `${sideNames[side]}可以完成直膝抬腿，但抬腿高度或膝蓋伸直程度未達最佳`);
  return result(1, `${sideNames[side]}直膝抬腿不足，可能抬腿高度不足或膝蓋彎曲明顯`);
}

function evaluatePushup(lm, pain) {
  if (pain) return result(0, "動作過程有疼痛，依 FMS 規則給 0 分");
  if (!visible(lm, [11, 12, 13, 14, 15, 16, 23, 24, 25, 26])) return result(1, "上肢與軀幹關節點辨識不清楚，無法完整判斷軀幹穩定伏地挺身");
  const angles = poseAngles(lm);
  const elbow = average([angles.left_elbow, angles.right_elbow]);
  const shoulders = midpoint(lm, 11, 12);
  const hips = midpoint(lm, 23, 24);
  const knees = midpoint(lm, 25, 26);
  const bodyLine = pointLineDistance(hips, shoulders, knees);
  const hipDrop = Math.abs(hips.y - ((shoulders.y + knees.y) / 2));
  if (elbow >= 155 && bodyLine <= 0.05 && hipDrop <= 0.06) return result(3, "手肘伸直時軀幹維持一直線，核心穩定度良好");
  if (elbow >= 135 && bodyLine <= 0.10) return result(2, "可以完成伏地挺身，但軀幹有下沉、抬臀或穩定度代償");
  return result(1, "無法完成穩定伏地挺身，可能手肘推起不足或軀幹無法維持穩定");
}

function evaluateRotary(lm, pain, side, phase) {
  if (pain) return result(0, `${sideNames[side]}動作過程有疼痛，依 FMS 規則給 0 分`);
  const [wrist, elbow, ankle, oppositeAnkle, oppositeKnee] = side === "left" ? [15, 13, 27, 28, 26] : [16, 14, 28, 27, 25];
  if (!visible(lm, [11, 12, 23, 24, wrist, elbow, ankle, oppositeAnkle, oppositeKnee])) return result(1, `${sideNames[side]}手腕、髖或踝關節點辨識不清楚`);
  const shoulderLevel = Math.abs(lm[11].y - lm[12].y);
  const pelvisLevel = Math.abs(lm[23].y - lm[24].y);
  if (phase === "same") {
    const extension = Math.abs(lm[wrist].x - lm[ankle].x);
    if (extension >= 0.22 && shoulderLevel <= 0.05 && pelvisLevel <= 0.05 && torsoLean(lm) <= 25) {
      return result(3, `${sideNames[side]}同側手腳伸展時軀幹穩定，肩膀與骨盆控制良好`);
    }
    return result(1, `${sideNames[side]}同側手腳伸展未達 3 分，請改做對側手肘碰膝蓋測驗`);
  }
  const elbowToKnee = dist(lm[elbow], lm[oppositeKnee]);
  const target = side === "left" ? "左手肘碰右膝蓋" : "右手肘碰左膝蓋";
  if (elbowToKnee <= 0.12 && shoulderLevel <= 0.08 && pelvisLevel <= 0.08) return result(2, `${target}可以完成，軀幹與骨盆控制尚可`);
  return result(1, `${target}未完成或軀幹骨盆晃動明顯，旋轉穩定性不足`);
}

function evaluateFms(lm) {
  const movement = currentMovement();
  if (movement.id === "deep_squat") return evaluateDeepSquat(lm, painReported);
  if (movement.id === "hurdle_step") return evaluateHurdleStep(lm, painReported, currentSide);
  if (movement.id === "inline_lunge") return evaluateInlineLunge(lm, painReported, currentSide);
  if (movement.id === "shoulder_mobility") return evaluateShoulderMobility(lm, painReported, currentSide);
  if (movement.id === "active_straight_leg_raise") return evaluateLegRaise(lm, painReported, currentSide);
  if (movement.id === "trunk_stability_pushup") return evaluatePushup(lm, painReported);
  return evaluateRotary(lm, painReported, currentSide, rotaryPhase);
}

function summarizeRecording() {
  if (!rows.length) return null;
  let finalScore;
  let reason;
  let bestRow = rows[rows.length - 1];
  let manual = false;
  if (painReported) {
    finalScore = 0;
    reason = "動作過程有疼痛，依 FMS 規則給 0 分";
  } else if (manualOverrideScore !== null) {
    finalScore = manualOverrideScore;
    reason = `${bestRow.reason}；人工覆寫為 ${manualOverrideScore} 分，以人工判定為主`;
    manual = true;
  } else {
    bestRow = rows.reduce((best, row) => row.score > best.score ? row : best, rows[0]);
    finalScore = bestRow.score;
    reason = bestRow.reason;
  }
  const duration = rows[rows.length - 1].elapsed;
  return {
    movement_id: currentMovement().id,
    movement_name: currentMovement().name,
    side: currentMovement().bilateral ? currentSide : "",
    final_score: finalScore,
    reason,
    samples: rows.length,
    duration,
    average_fps: duration ? Math.round(rows.length / duration * 1000) / 1000 : 0,
    best_frame_score: manual ? finalScore : bestRow.score,
    auto_score: bestRow.score,
    manual_override: manual,
  };
}

function combineBilateral(movement) {
  const sides = sessionSideResults[movement.id] || {};
  if (!sides.left || !sides.right) return null;
  const lower = sides.left.final_score <= sides.right.final_score ? sides.left : sides.right;
  const samples = sides.left.samples + sides.right.samples;
  const duration = Math.round((sides.left.duration + sides.right.duration) * 1000000) / 1000000;
  return {
    movement_id: movement.id,
    movement_name: movement.name,
    final_score: lower.final_score,
    left_score: sides.left.final_score,
    right_score: sides.right.final_score,
    reason: `左側 ${sides.left.final_score} 分：${sides.left.reason}；右側 ${sides.right.final_score} 分：${sides.right.reason}。最後取較低分的 ${sideNames[lower.side]}。`,
    samples,
    duration,
    average_fps: duration ? Math.round(samples / duration * 1000) / 1000 : 0,
    best_frame_score: lower.best_frame_score,
    auto_score: lower.auto_score,
    manual_override: sides.left.manual_override || sides.right.manual_override,
  };
}

function chooseNextMovement() {
  for (let i = 0; i < movements.length; i += 1) {
    if (!sessionResults[movements[i].id]) {
      movementIndex = i;
      currentSide = "left";
      rotaryPhase = "same";
      manualOverrideScore = null;
      return;
    }
  }
}

function stopRecording() {
  recording = false;
  recordBtn.textContent = "開始";
  const summary = summarizeRecording();
  if (!summary) {
    setStatus("沒有擷取到有效資料");
    return;
  }
  const movement = currentMovement();
  if (movement.id === "rotary_stability" && rotaryPhase === "same" && summary.final_score < 3 && !summary.manual_override) {
    rotaryPhase = "diagonal";
    rows = [];
    manualOverrideScore = null;
    setStatus(`${sideNames[currentSide]}同側未達 3 分，請做對側手肘碰膝蓋`);
    updateUi();
    return;
  }
  if (movement.bilateral) {
    sessionSideResults[movement.id] ||= {};
    sessionSideResults[movement.id][currentSide] = summary;
    const combined = combineBilateral(movement);
    if (combined) sessionResults[movement.id] = combined;
  } else {
    sessionResults[movement.id] = { ...summary, left_score: "", right_score: "" };
  }
  if (Object.keys(sessionResults).length === movements.length) {
    setStatus("七個動作完成，請輸入測驗名稱");
    saveDialog.showModal();
  } else if (movement.bilateral && !sessionResults[movement.id]) {
    currentSide = currentSide === "left" ? "right" : "left";
    rotaryPhase = "same";
    manualOverrideScore = null;
    setStatus(`${movement.name} 已完成 ${Object.keys(sessionSideResults[movement.id] || {}).length}/2 側`);
  } else {
    chooseNextMovement();
    setStatus(`已完成 ${Object.keys(sessionResults).length}/7 動作`);
  }
  updateUi();
}

function startRecording() {
  recording = true;
  rows = [];
  recordStart = performance.now();
  sampleCount = 0;
  painReported = false;
  recordBtn.textContent = "結束";
  setStatus("錄製中...");
  updateUi();
}

function toggleRecording() {
  if (recording) stopRecording();
  else startRecording();
}

function sanitizeName(name) {
  return (name || "FMS測驗").trim().replace(/[\\/:*?"<>|]/g, "_") || "FMS測驗";
}

function resetForNextTest(statusText = "已準備下一位受測者") {
  recording = false;
  rows = [];
  recordStart = 0;
  sampleCount = 0;
  painReported = false;
  manualOverrideScore = null;
  movementIndex = 0;
  currentSide = "left";
  rotaryPhase = "same";
  sessionResults = {};
  sessionSideResults = {};
  recordBtn.textContent = "開始";
  scoreTextEl.textContent = "-";
  reasonTextEl.textContent = "";
  testNameInput.value = "";
  if (downloadDialog.open) downloadDialog.close();
  if (pendingDownloadUrl) {
    URL.revokeObjectURL(pendingDownloadUrl);
    pendingDownloadUrl = null;
  }
  pendingExcelBlob = null;
  pendingExcelFileName = "";
  cloudUploadBtn.disabled = false;
  cloudUploadBtn.textContent = "上傳到 Google Drive";
  setStatus(statusText);
  updateUi();
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const value = String(reader.result || "");
      resolve(value.includes(",") ? value.split(",")[1] : value);
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

async function uploadExcelToGoogleDrive() {
  if (!pendingExcelBlob || !pendingExcelFileName) {
    setStatus("目前沒有可上傳的 Excel");
    return;
  }
  if (!cloudFolderName) {
    folderDialog.showModal();
    setStatus("請先建立或選擇測驗資料夾");
    return;
  }
  cloudUploadBtn.disabled = true;
  cloudUploadBtn.textContent = "上傳中...";
  setStatus(`正在上傳到 Google Drive：${cloudFolderName}`);
  const base64 = await blobToBase64(pendingExcelBlob);
  const formData = new FormData();
  formData.append("action", "upload");
  formData.append("folderName", cloudFolderName);
  formData.append("fileName", pendingExcelFileName);
  formData.append("mimeType", pendingExcelBlob.type);
  formData.append("base64", base64);

  try {
    await fetch(GOOGLE_UPLOAD_URL, {
      method: "POST",
      body: formData,
      mode: "no-cors",
    });
    downloadInfo.textContent = `已送出上傳請求，檔案會放在 Google Drive 的「${cloudFolderName}」資料夾。`;
    setStatus(`已送出上傳：${cloudFolderName}`);
    cloudUploadBtn.textContent = "已送出上傳";
    setTimeout(() => resetForNextTest("已上傳，請開始下一位受測者"), 900);
  } catch (error) {
    cloudUploadBtn.disabled = false;
    cloudUploadBtn.textContent = "重新上傳到 Google Drive";
    throw error;
  }
}

function createRadarChartDataUrl(scoreRows, total) {
  const size = 900;
  const canvasEl = document.createElement("canvas");
  canvasEl.width = size;
  canvasEl.height = size;
  const chartCtx = canvasEl.getContext("2d");
  const centerX = size / 2;
  const centerY = 500;
  const radius = 250;
  const maxScore = 3;
  const labels = scoreRows.map((row) => row.動作名稱);
  const scores = scoreRows.map((row) => Number(row.最後分數) || 0);

  chartCtx.fillStyle = "#ffffff";
  chartCtx.fillRect(0, 0, size, size);
  chartCtx.fillStyle = "#111827";
  chartCtx.font = "700 34px -apple-system, BlinkMacSystemFont, 'PingFang TC', sans-serif";
  chartCtx.textAlign = "center";
  chartCtx.fillText("FMS 七項分數雷達圖", centerX, 64);
  chartCtx.font = "24px -apple-system, BlinkMacSystemFont, 'PingFang TC', sans-serif";
  chartCtx.fillText(`總分：${total} / 21`, centerX, 104);

  for (let level = 1; level <= maxScore; level += 1) {
    const levelRadius = radius * (level / maxScore);
    chartCtx.beginPath();
    labels.forEach((_, index) => {
      const angle = -Math.PI / 2 + index * 2 * Math.PI / labels.length;
      const x = centerX + Math.cos(angle) * levelRadius;
      const y = centerY + Math.sin(angle) * levelRadius;
      if (index === 0) chartCtx.moveTo(x, y);
      else chartCtx.lineTo(x, y);
    });
    chartCtx.closePath();
    chartCtx.strokeStyle = level === maxScore ? "#6b7280" : "#d1d5db";
    chartCtx.lineWidth = level === maxScore ? 2 : 1;
    chartCtx.stroke();
  }

  labels.forEach((label, index) => {
    const angle = -Math.PI / 2 + index * 2 * Math.PI / labels.length;
    const axisX = centerX + Math.cos(angle) * radius;
    const axisY = centerY + Math.sin(angle) * radius;
    const labelX = centerX + Math.cos(angle) * (radius + 86);
    const labelY = centerY + Math.sin(angle) * (radius + 66);
    chartCtx.beginPath();
    chartCtx.moveTo(centerX, centerY);
    chartCtx.lineTo(axisX, axisY);
    chartCtx.strokeStyle = "#e5e7eb";
    chartCtx.lineWidth = 1;
    chartCtx.stroke();

    chartCtx.fillStyle = "#111827";
    chartCtx.font = "22px -apple-system, BlinkMacSystemFont, 'PingFang TC', sans-serif";
    chartCtx.textAlign = labelX < centerX - 20 ? "right" : labelX > centerX + 20 ? "left" : "center";
    chartCtx.fillText(label, labelX, labelY);
    chartCtx.font = "700 24px -apple-system, BlinkMacSystemFont, 'PingFang TC', sans-serif";
    chartCtx.fillText(`${scores[index]}分`, labelX, labelY + 30);
  });

  chartCtx.beginPath();
  scores.forEach((score, index) => {
    const angle = -Math.PI / 2 + index * 2 * Math.PI / scores.length;
    const pointRadius = radius * (score / maxScore);
    const x = centerX + Math.cos(angle) * pointRadius;
    const y = centerY + Math.sin(angle) * pointRadius;
    if (index === 0) chartCtx.moveTo(x, y);
    else chartCtx.lineTo(x, y);
  });
  chartCtx.closePath();
  chartCtx.fillStyle = "rgba(8, 127, 91, 0.24)";
  chartCtx.strokeStyle = "#087f5b";
  chartCtx.lineWidth = 5;
  chartCtx.fill();
  chartCtx.stroke();

  scores.forEach((score, index) => {
    const angle = -Math.PI / 2 + index * 2 * Math.PI / scores.length;
    const pointRadius = radius * (score / maxScore);
    const x = centerX + Math.cos(angle) * pointRadius;
    const y = centerY + Math.sin(angle) * pointRadius;
    chartCtx.beginPath();
    chartCtx.arc(x, y, 7, 0, 2 * Math.PI);
    chartCtx.fillStyle = "#087f5b";
    chartCtx.fill();
    chartCtx.strokeStyle = "#ffffff";
    chartCtx.lineWidth = 3;
    chartCtx.stroke();
  });

  return canvasEl.toDataURL("image/png");
}

async function saveWorkbook(testName) {
  const rowsOut = [];
  let total = 0;
  for (const movement of movements) {
    const item = sessionResults[movement.id];
    total += item.final_score;
    rowsOut.push({
      測驗資料夾: cloudFolderName || "未設定",
      測驗名稱: testName,
      動作名稱: item.movement_name,
      最後分數: item.final_score,
      左側分數: item.left_score ?? "",
      右側分數: item.right_score ?? "",
      原因: item.reason,
      人工覆寫: item.manual_override ? "是" : "否",
      自動辨識分數: item.auto_score,
      總擷取筆數: item.samples,
      測試秒數: item.duration,
      平均擷取FPS: item.average_fps,
      採用分數: item.best_frame_score,
      FMS總分: "",
    });
  }
  rowsOut.push({ 測驗資料夾: cloudFolderName || "未設定", 測驗名稱: testName, 動作名稱: "總分", 原因: "七個 FMS 動作分數加總", FMS總分: total });
  const scoreRows = rowsOut.slice(0, movements.length);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "FMS 功能性動作篩檢";
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet("FMS分數");
  worksheet.columns = [
    { header: "測驗資料夾", key: "測驗資料夾", width: 22 },
    { header: "測驗名稱", key: "測驗名稱", width: 22 },
    { header: "動作名稱", key: "動作名稱", width: 18 },
    { header: "最後分數", key: "最後分數", width: 10 },
    { header: "左側分數", key: "左側分數", width: 10 },
    { header: "右側分數", key: "右側分數", width: 10 },
    { header: "原因", key: "原因", width: 70 },
    { header: "人工覆寫", key: "人工覆寫", width: 10 },
    { header: "自動辨識分數", key: "自動辨識分數", width: 14 },
    { header: "總擷取筆數", key: "總擷取筆數", width: 12 },
    { header: "測試秒數", key: "測試秒數", width: 12 },
    { header: "平均擷取FPS", key: "平均擷取FPS", width: 14 },
    { header: "採用分數", key: "採用分數", width: 10 },
    { header: "FMS總分", key: "FMS總分", width: 10 },
  ];
  rowsOut.forEach((row) => worksheet.addRow(row));
  worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  worksheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF087F5B" } };
  worksheet.getColumn("原因").alignment = { wrapText: true, vertical: "top" };
  worksheet.eachRow((row) => {
    row.alignment = { vertical: "top" };
  });

  const radarSheet = workbook.addWorksheet("雷達圖");
  radarSheet.getCell("A1").value = "FMS 七項分數雷達圖";
  radarSheet.getCell("A1").font = { bold: true, size: 18 };
  radarSheet.getCell("A2").value = `測驗名稱：${testName}`;
  radarSheet.getCell("A3").value = `測驗資料夾：${cloudFolderName || "未設定"}`;
  radarSheet.getCell("A4").value = `總分：${total} / 21`;
  radarSheet.addRow([]);
  radarSheet.addRow(["動作名稱", "分數"]);
  scoreRows.forEach((row) => radarSheet.addRow([row.動作名稱, row.最後分數]));
  radarSheet.getColumn(1).width = 22;
  radarSheet.getColumn(2).width = 10;
  radarSheet.getRow(6).font = { bold: true };
  const radarImage = workbook.addImage({
    base64: createRadarChartDataUrl(scoreRows, total),
    extension: "png",
  });
  radarSheet.addImage(radarImage, {
    tl: { col: 3, row: 0 },
    ext: { width: 720, height: 720 },
  });

  const fileName = `${sanitizeName(testName)}_${new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-")}.xlsx`;
  const data = await workbook.xlsx.writeBuffer();
  const blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  pendingExcelBlob = blob;
  pendingExcelFileName = fileName;
  cloudUploadBtn.disabled = false;
  cloudUploadBtn.textContent = "上傳到 Google Drive";
  if ("showDirectoryPicker" in window) {
    const dir = await window.showDirectoryPicker();
    const handle = await dir.getFileHandle(fileName, { create: true });
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    setStatus("Excel 已儲存");
  } else {
    if (pendingDownloadUrl) URL.revokeObjectURL(pendingDownloadUrl);
    const url = URL.createObjectURL(blob);
    pendingDownloadUrl = url;
    downloadLink.href = url;
    downloadLink.download = fileName;
    downloadLink.textContent = `下載 ${fileName}`;
    downloadInfo.textContent = "可以下載到手機，也可以直接上傳到 Google Drive。";
    downloadDialog.showModal();
    setStatus("Excel 已產生，請按下載");
  }
}

function updateUi() {
  const movement = currentMovement();
  movementNameEl.textContent = movement.name;
  sideRowEl.hidden = !movement.bilateral;
  sideButtons.hidden = !movement.bilateral;
  sideNameEl.textContent = sideNames[currentSide];
  rotaryRowEl.hidden = movement.id !== "rotary_stability";
  rotaryPhaseEl.textContent = rotaryPhase === "same" ? "同側手腳伸展" : "對側手肘碰膝蓋";
  completedText.textContent = `${Object.keys(sessionResults).length}/7`;
  samplesText.textContent = String(sampleCount);
  fpsText.textContent = measuredFps.toFixed(1);
  painBtn.classList.toggle("danger", painReported);
  document.querySelectorAll(".manual button").forEach((button) => {
    button.classList.toggle("active", Number(button.dataset.score) === manualOverrideScore);
  });
  sideButtons.querySelectorAll("button").forEach((button) => {
    const side = button.dataset.side;
    button.classList.toggle("active", side === currentSide);
    button.classList.toggle("done", Boolean(sessionSideResults[movement.id]?.[side]));
  });
  movementList.querySelectorAll("button").forEach((button, index) => {
    button.classList.toggle("active", index === movementIndex);
    button.classList.toggle("done", Boolean(sessionResults[movements[index].id]));
  });
}

function renderMovementButtons() {
  movementList.innerHTML = "";
  movements.forEach((movement, index) => {
    const button = document.createElement("button");
    button.textContent = movement.name;
    button.addEventListener("click", () => {
      if (recording) return;
      movementIndex = index;
      currentSide = "left";
      rotaryPhase = "same";
      manualOverrideScore = null;
      scoreTextEl.textContent = "-";
      reasonTextEl.textContent = "";
      updateUi();
    });
    movementList.appendChild(button);
  });
}

function drawResults(results) {
  const width = canvas.width;
  const height = canvas.height;
  ctx.clearRect(0, 0, width, height);
  ctx.save();
  ctx.scale(-1, 1);
  ctx.translate(-width, 0);
  drawConnectors(ctx, results.poseLandmarks, POSE_CONNECTIONS, { color: "#4ade80", lineWidth: 3 });
  drawLandmarks(ctx, results.poseLandmarks, { color: "#fbbf24", lineWidth: 1, radius: 3 });
  ctx.restore();
}

function onResults(results) {
  if (video.videoWidth && (canvas.width !== video.videoWidth || canvas.height !== video.videoHeight)) {
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
  }
  const now = performance.now();
  measuredFps = 1000 / Math.max(1, now - lastFrameTime);
  lastFrameTime = now;
  if (results.poseLandmarks) {
    lastPoseLandmarks = results.poseLandmarks;
    drawResults(results);
    const fms = evaluateFms(results.poseLandmarks);
    const displayScore = manualOverrideScore ?? fms.score;
    scoreTextEl.textContent = String(displayScore);
    reasonTextEl.textContent = manualOverrideScore === null ? fms.reason : `已選擇人工覆寫：${manualOverrideScore} 分`;
    if (recording) {
      sampleCount += 1;
      rows.push({
        movement: currentMovement().name,
        side: currentMovement().bilateral ? currentSide : "",
        score: fms.score,
        reason: fms.reason,
        sample: sampleCount,
        elapsed: Math.round((performance.now() - recordStart) / 1000 * 1000000) / 1000000,
        fps: Math.round(measuredFps * 1000) / 1000,
      });
    }
  } else {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
  }
  updateUi();
}

async function processLoop() {
  if (cameraReady && !processing) {
    processing = true;
    await pose.send({ image: video });
    processing = false;
  }
  requestAnimationFrame(processLoop);
}

async function openCamera() {
  if (cameraStream) {
    cameraStream.getTracks().forEach((track) => track.stop());
  }
  cameraStream = await navigator.mediaDevices.getUserMedia({
    video: {
      facingMode: cameraFacingMode,
      width: { ideal: 1280 },
      height: { ideal: 720 },
      frameRate: { ideal: 60 },
    },
    audio: false,
  });
  video.srcObject = cameraStream;
  await video.play();
  video.style.transform = cameraFacingMode === "user" ? "scaleX(-1)" : "none";
  await configureCameraZoom();
  cameraReady = true;
  setStatus(cameraFacingMode === "user" ? "前鏡頭已啟動" : "後鏡頭已啟動");
}

async function startCamera() {
  pose = new Pose({
    locateFile: (file) => `https://cdn.jsdelivr.net/npm/@mediapipe/pose/${file}`,
  });
  pose.setOptions({
    modelComplexity: 1,
    smoothLandmarks: true,
    enableSegmentation: false,
    minDetectionConfidence: 0.7,
    minTrackingConfidence: 0.7,
  });
  pose.onResults(onResults);

  await openCamera();
  requestAnimationFrame(processLoop);
}

recordBtn.addEventListener("click", toggleRecording);
painBtn.addEventListener("click", () => {
  painReported = !painReported;
  updateUi();
});
changeFolderBtn.addEventListener("click", () => {
  folderNameInput.value = cloudFolderName;
  folderDialog.showModal();
});
folderStartBtn.addEventListener("click", async (event) => {
  event.preventDefault();
  folderStartBtn.disabled = true;
  folderStartBtn.textContent = "建立中...";
  try {
    await enterCloudFolder(folderNameInput.value || "FMS測驗");
  } finally {
    folderStartBtn.disabled = false;
    folderStartBtn.textContent = "建立並進入";
  }
});
switchCameraBtn.addEventListener("click", async () => {
  cameraFacingMode = cameraFacingMode === "user" ? "environment" : "user";
  cameraReady = false;
  setStatus("正在切換鏡頭...");
  try {
    await openCamera();
  } catch (error) {
    console.error(error);
    cameraFacingMode = "user";
    setStatus("切換鏡頭失敗，請確認瀏覽器相機權限");
  }
});
zoomSlider.addEventListener("input", async () => {
  try {
    await applyCameraZoom(zoomSlider.value);
  } catch (error) {
    console.error(error);
    setStatus("鏡頭遠近調整失敗");
  }
});
sideButtons.addEventListener("click", (event) => {
  if (recording || !event.target.dataset.side) return;
  currentSide = event.target.dataset.side;
  rotaryPhase = "same";
  manualOverrideScore = null;
  updateUi();
});
document.querySelector(".manual").addEventListener("click", (event) => {
  if (!event.target.dataset.score) return;
  const score = Number(event.target.dataset.score);
  manualOverrideScore = manualOverrideScore === score ? null : score;
  updateUi();
});
saveBtn.addEventListener("click", async (event) => {
  event.preventDefault();
  const testName = sanitizeName(testNameInput.value);
  saveDialog.close();
  try {
    await saveWorkbook(testName);
  } catch (error) {
    console.error(error);
    setStatus("儲存失敗或已取消");
    return;
  }
});
downloadLink.addEventListener("click", () => {
  setStatus("下載已送出，請到手機的下載項目查看");
});
cloudUploadBtn.addEventListener("click", async () => {
  try {
    await uploadExcelToGoogleDrive();
  } catch (error) {
    console.error(error);
    cloudUploadBtn.disabled = false;
    cloudUploadBtn.textContent = "重新上傳到 Google Drive";
    downloadInfo.textContent = "上傳失敗，請確認網路或 Google Apps Script 權限，仍可先下載 Excel 備份。";
    setStatus("Google Drive 上傳失敗");
  }
});
finishBtn.addEventListener("click", () => {
  resetForNextTest("測驗完成，請開始下一位受測者");
});

renderMovementButtons();
updateFolderUi();
updateUi();
if (!cloudFolderName) {
  setTimeout(() => folderDialog.showModal(), 300);
}
startCamera().catch((error) => {
  console.error(error);
  setStatus("無法啟動相機，請確認瀏覽器權限與 HTTPS/localhost");
});
