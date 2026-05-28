const passwordInput = document.querySelector("#password");
const nameInput = document.querySelector("#nameContext");
const emailInput = document.querySelector("#emailContext");
const mobileInput = document.querySelector("#mobileContext");
const toggleButton = document.querySelector("#togglePassword");
const toggleIcon = document.querySelector("#toggleIcon");
const copyButton = document.querySelector("#copyPassword");
const copyIcon = document.querySelector("#copyIcon");
const generateButton = document.querySelector("#generatePassword");
const passwordLengthInput = document.querySelector("#passwordLength");

const scoreRing = document.querySelector(".score-ring");
const scoreValue = document.querySelector("#scoreValue");
const scoreLabel = document.querySelector("#scoreLabel");
const analysisStatus = document.querySelector("#analysisStatus");
const modelVersion = document.querySelector("#modelVersion");
const aiProbability = document.querySelector("#aiProbability");
const aiPrediction = document.querySelector("#aiPrediction");
const entropyBits = document.querySelector("#entropyBits");
const effectiveEntropy = document.querySelector("#effectiveEntropy");
const offlineTime = document.querySelector("#offlineTime");
const onlineTime = document.querySelector("#onlineTime");
const strengthList = document.querySelector("#strengthList");
const riskList = document.querySelector("#riskList");
const suggestionList = document.querySelector("#suggestionList");
const explanationList = document.querySelector("#explanationList");

let debounceHandle;
let currentScore = 0;
let scoreAnimationFrame;
let copyStatusHandle;
let analysisRequestId = 0;
const minimumLoadingTime = 350;

function contextPayload() {
  return {
    name: nameInput.value,
    email: emailInput.value,
    mobile: mobileInput.value,
  };
}

function setList(element, items) {
  element.innerHTML = "";
  items.forEach((item) => {
    const li = document.createElement("li");
    li.textContent = item;
    element.appendChild(li);
  });
}

function updateCopyButton() {
  copyButton.disabled = passwordInput.value.length === 0;
}

function ringColor(score) {
  if (score <= 0) return "#cfd8e3";
  if (score < 45) return "#b42318";
  if (score < 65) return "#f79009";
  return "#18815f";
}

function render(result) {
  const score = result.score ?? 0;
  animateScore(score);
  scoreRing.style.setProperty("--ring-color", ringColor(score));
  scoreLabel.textContent = result.label ?? "Empty";
  modelVersion.textContent = result.model_version ?? "local";
  aiProbability.textContent = `${Math.round((result.ai_confidence ?? result.ai_probability ?? 0) * 100)}%`;
  aiPrediction.textContent = result.ai_prediction ?? "Waiting";
  entropyBits.textContent = `${result.entropy_bits ?? 0} bits`;
  effectiveEntropy.textContent = `${result.effective_entropy_bits ?? 0} effective bits`;
  offlineTime.textContent = result.estimated_crack_time?.offline_fast_hash ?? "Instantly";
  onlineTime.textContent = result.estimated_crack_time?.online_throttled ?? "Instantly";
  setList(strengthList, result.strengths ?? []);
  setList(riskList, result.risks ?? []);
  setList(suggestionList, result.suggestions ?? []);
  setList(explanationList, result.ai_explanation ?? []);
}

function animateScore(targetScore) {
  if (scoreAnimationFrame) {
    cancelAnimationFrame(scoreAnimationFrame);
  }

  const startScore = currentScore;
  const distance = targetScore - startScore;
  const duration = 420;
  const startedAt = performance.now();

  function step(now) {
    const progress = Math.min((now - startedAt) / duration, 1);
    const eased = 1 - Math.pow(1 - progress, 3);
    currentScore = Math.round(startScore + distance * eased);
    scoreRing.style.setProperty("--score", currentScore);
    scoreValue.textContent = currentScore;

    if (progress < 1) {
      scoreAnimationFrame = requestAnimationFrame(step);
    } else {
      currentScore = targetScore;
      scoreRing.style.setProperty("--score", targetScore);
      scoreValue.textContent = targetScore;
      scoreAnimationFrame = null;
    }
  }

  scoreAnimationFrame = requestAnimationFrame(step);
}

async function analyze() {
  const requestId = ++analysisRequestId;
  const startedAt = performance.now();
  setAnalyzing(true);

  try {
    const response = await fetch("/api/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        password: passwordInput.value,
        context: contextPayload(),
      }),
    });

    if (!response.ok) {
      throw new Error("Analysis request failed");
    }

    const result = await response.json();
    if (requestId === analysisRequestId) {
      render(result);
    }
  } finally {
    if (requestId === analysisRequestId) {
      await keepLoadingVisible(startedAt);
      setAnalyzing(false);
    }
  }
}

async function generatePassword() {
  const startedAt = performance.now();
  generateButton.disabled = true;
  generateButton.textContent = "Generating...";
  setAnalyzing(true);

  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        length: Number(passwordLengthInput.value || 18),
      }),
    });

    if (!response.ok) {
      throw new Error("Password generation failed");
    }

    const result = await response.json();
    passwordInput.value = result.password;
    updateCopyButton();
    render(result.analysis);
  } catch {
    render({
      score: 0,
      label: "Generator Error",
      ai_probability: 0,
      ai_confidence: 0,
      ai_prediction: "Unavailable",
      entropy_bits: 0,
      effective_entropy_bits: 0,
      estimated_crack_time: {
        offline_fast_hash: "Unavailable",
        online_throttled: "Unavailable",
      },
      strengths: [],
      risks: ["Password generator could not be reached."],
      suggestions: ["Restart the Python server and try again."],
      ai_explanation: ["The generator endpoint did not return a password."],
      model_version: "not connected",
    });
  } finally {
    await keepLoadingVisible(startedAt);
    generateButton.disabled = false;
    generateButton.textContent = "Generate Strong Password";
    setAnalyzing(false);
  }
}

async function copyPassword() {
  const password = passwordInput.value;
  if (!password) {
    setCopyStatus("Empty");
    return;
  }

  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(password);
    } else {
      fallbackCopy(password);
    }
    setCopyStatus("Copied");
  } catch {
    const copied = fallbackCopy(password);
    setCopyStatus(copied ? "Copied" : "Failed");
  }
}

function fallbackCopy(value) {
  const textArea = document.createElement("textarea");
  textArea.value = value;
  textArea.setAttribute("readonly", "");
  textArea.style.position = "fixed";
  textArea.style.left = "-9999px";
  document.body.appendChild(textArea);
  textArea.select();

  let copied = false;
  try {
    copied = document.execCommand("copy");
  } finally {
    document.body.removeChild(textArea);
  }
  return copied;
}

function setCopyStatus(status) {
  window.clearTimeout(copyStatusHandle);
  copyIcon.textContent = status;
  copyButton.classList.toggle("is-copied", status === "Copied");

  copyStatusHandle = window.setTimeout(() => {
    copyIcon.textContent = "Copy";
    copyButton.classList.remove("is-copied");
    updateCopyButton();
  }, 1200);
}

function setAnalyzing(isAnalyzing) {
  analysisStatus.hidden = !isAnalyzing;
}

async function keepLoadingVisible(startedAt) {
  const elapsed = performance.now() - startedAt;
  const remaining = minimumLoadingTime - elapsed;
  if (remaining > 0) {
    await new Promise((resolve) => setTimeout(resolve, remaining));
  }
}

function scheduleAnalyze() {
  window.clearTimeout(debounceHandle);
  debounceHandle = window.setTimeout(() => {
    analyze().catch(() => {
      render({
        score: 0,
        label: "Offline",
        ai_probability: 0,
        ai_confidence: 0,
        ai_prediction: "Offline",
        entropy_bits: 0,
        effective_entropy_bits: 0,
        estimated_crack_time: {
          offline_fast_hash: "Unavailable",
          online_throttled: "Unavailable",
        },
        strengths: [],
        risks: ["Start the Python server to use the analyzer."],
        suggestions: ["Run python app.py from the project folder."],
        ai_explanation: ["The backend API is offline, so the AI explanation is unavailable."],
        model_version: "not connected",
      });
    });
  }, 180);
}

toggleButton.addEventListener("click", () => {
  const isPassword = passwordInput.type === "password";
  passwordInput.type = isPassword ? "text" : "password";
  toggleIcon.textContent = isPassword ? "Hide" : "Show";
  toggleButton.setAttribute("aria-label", isPassword ? "Hide password" : "Show password");
  toggleButton.setAttribute("title", isPassword ? "Hide password" : "Show password");
});

generateButton.addEventListener("click", generatePassword);
copyButton.addEventListener("click", copyPassword);

[passwordInput, nameInput, emailInput, mobileInput].forEach((element) => {
  element.addEventListener("input", scheduleAnalyze);
});

passwordInput.addEventListener("input", updateCopyButton);

updateCopyButton();
scheduleAnalyze();
