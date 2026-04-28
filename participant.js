const GOOGLE_UPLOAD_URL = "https://script.google.com/macros/s/AKfycbxQ9cKAp--d4CV_Vf0NxVXVaz2t8MBpxKSPaBDtvTmsH14HIaWeCTIrEBaJ8ScIT5cEsg/exec";

const statusEl = document.getElementById("participantStatus");
const form = document.getElementById("participantForm");
const sessionSelect = document.getElementById("sessionSelect");
const nameInput = document.getElementById("participantName");
const genderInput = document.getElementById("participantGender");
const ageInput = document.getElementById("participantAge");
const exerciseInput = document.getElementById("participantExercise");
const initialFolder = new URLSearchParams(window.location.search).get("folder") || "";

function setStatus(text) {
  statusEl.textContent = text;
}

function sanitizeName(name) {
  return String(name || "FMS測驗").trim().replace(/[\\/:*?"<>|]/g, "_") || "FMS測驗";
}

function jsonp(action, params = {}) {
  return new Promise((resolve, reject) => {
    const callback = `fmsCallback_${Date.now()}_${Math.round(Math.random() * 100000)}`;
    const url = new URL(GOOGLE_UPLOAD_URL);
    url.searchParams.set("action", action);
    url.searchParams.set("callback", callback);
    Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));
    const script = document.createElement("script");
    const timer = setTimeout(() => {
      cleanup();
      reject(new Error("讀取資料逾時"));
    }, 12000);
    function cleanup() {
      clearTimeout(timer);
      script.remove();
      delete window[callback];
    }
    window[callback] = (data) => {
      cleanup();
      resolve(data);
    };
    script.onerror = () => {
      cleanup();
      reject(new Error("讀取資料失敗"));
    };
    script.src = url.toString();
    document.body.appendChild(script);
  });
}

async function loadSessions() {
  try {
    const data = await jsonp("listFolders");
    const folders = data.folders || [];
    sessionSelect.innerHTML = "";
    if (!folders.length) {
      sessionSelect.innerHTML = '<option value="">目前沒有檢測名稱</option>';
      setStatus("請先請檢測人員建立檢測名稱");
      return;
    }
    sessionSelect.append(new Option("請選擇檢測名稱", ""));
    folders.forEach((folder) => {
      sessionSelect.append(new Option(folder.name, folder.name));
    });
    if (initialFolder) {
      sessionSelect.value = initialFolder;
    }
    setStatus("請填寫基本資料");
  } catch (error) {
    console.error(error);
    sessionSelect.innerHTML = '<option value="">載入失敗</option>';
    setStatus("載入檢測名稱失敗，請稍後再試");
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const folderName = sessionSelect.value;
  const subjectName = sanitizeName(nameInput.value);
  if (!folderName || !subjectName) return;
  const formData = new FormData();
  formData.append("action", "createSubject");
  formData.append("folderName", folderName);
  formData.append("subjectName", subjectName);
  formData.append("gender", genderInput.value);
  formData.append("age", ageInput.value);
  formData.append("exerciseHabit", exerciseInput.value);

  setStatus("正在送出...");
  try {
    await fetch(GOOGLE_UPLOAD_URL, {
      method: "POST",
      body: formData,
      mode: "no-cors",
    });
    setStatus("已送出，請告知檢測人員更新名單");
    form.reset();
  } catch (error) {
    console.error(error);
    setStatus("送出失敗，請確認網路後再試一次");
  }
});

loadSessions();
