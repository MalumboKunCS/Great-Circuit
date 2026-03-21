// Apps Script Web App URL (Deploy → Web app → copy URL ending in /exec)
const APPS_SCRIPT_WEB_APP_URL =
  "https://script.google.com/macros/s/AKfycbwF43qtt9qneOpZblT5Jx89DI-gjpRLNgZ0WZM8_hn0YtMRNlceGDgPwiTU9k8-ia1dSA/exec";

/** After compression, keep uploads small so Apps Script receives the full POST (large base64 was dropping rows). */
/** Original file from device (we compress before upload). */
const MAX_PAYMENT_IMAGE_BYTES = 8 * 1024 * 1024;
const TARGET_MAX_COMPRESSED_BYTES = 450 * 1024;
const SUCCESS_TOAST_MS = 22000;

const SUCCESS_MESSAGE =
  "Thank you for registering. Your submission has been received and is subject to verification. Once your registration is confirmed, your digital ticket will be sent to the email address you provided, typically within 24–48 hours. We appreciate your patience and look forward to hosting you.";

const HEARD_THROUGH_PERSON = "Through a person";

function $(id) {
  return document.getElementById(id);
}

function trimValue(v) {
  return (v ?? "").toString().trim();
}

function isEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function showErrors(list) {
  const box = $("formErrors");
  const ul = $("errorList");
  ul.innerHTML = "";

  list.forEach((msg) => {
    const li = document.createElement("li");
    li.textContent = msg;
    ul.appendChild(li);
  });

  box.classList.remove("is-hidden");
  box.classList.add("global-toast--visible");
  $("formSuccess").classList.add("is-hidden");
  $("formSuccess").classList.remove("global-toast--visible");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function hideErrors() {
  const box = $("formErrors");
  box.classList.add("is-hidden");
  box.classList.remove("global-toast--visible");
}

let successToastTimer = null;

function showSuccess(msg) {
  const wrap = $("formSuccess");
  const textEl = $("formSuccessMessage");
  if (textEl) textEl.textContent = msg;
  wrap.classList.remove("is-hidden");
  wrap.classList.add("global-toast--visible");
  hideErrors();
  window.scrollTo({ top: 0, behavior: "smooth" });
  if (successToastTimer) clearTimeout(successToastTimer);
  successToastTimer = setTimeout(() => {
    hideSuccess();
  }, SUCCESS_TOAST_MS);
}

function hideSuccess() {
  const wrap = $("formSuccess");
  wrap.classList.add("is-hidden");
  wrap.classList.remove("global-toast--visible");
  if (successToastTimer) {
    clearTimeout(successToastTimer);
    successToastTimer = null;
  }
}

/**
 * Resize + JPEG compress so the POST payload stays small (Apps Script often truncates huge fields).
 */
function compressImageFile(file, maxEdge = 1400, quality = 0.82) {
  return new Promise((resolve, reject) => {
    if (!file.type.startsWith("image/")) {
      resolve(file);
      return;
    }
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      let w = img.naturalWidth || img.width;
      let h = img.naturalHeight || img.height;
      if (!w || !h) {
        reject(new Error("Invalid image"));
        return;
      }
      if (w > maxEdge || h > maxEdge) {
        if (w > h) {
          h = Math.round((h * maxEdge) / w);
          w = maxEdge;
        } else {
          w = Math.round((w * maxEdge) / h);
          h = maxEdge;
        }
      }
      const canvas = document.createElement("canvas");
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(img, 0, 0, w, h);
      canvas.toBlob(
        (blob) => {
          if (!blob) reject(new Error("Compression failed"));
          else resolve(blob);
        },
        "image/jpeg",
        quality
      );
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Could not load image"));
    };
    img.src = url;
  });
}

async function preparePaymentProofBlob(file) {
  if (!file.type.startsWith("image/")) return file;
  try {
    const blob = await compressImageFile(file);
    if (blob.size <= TARGET_MAX_COMPRESSED_BYTES) return blob;
    let q = 0.72;
    for (let attempt = 0; attempt < 4; attempt++) {
      const b = await compressImageFile(file, 1200, q);
      if (b.size <= TARGET_MAX_COMPRESSED_BYTES) return b;
      q -= 0.08;
    }
    return blob;
  } catch {
    if (file.size > MAX_PAYMENT_IMAGE_BYTES) throw new Error("Image too large");
    return file;
  }
}

function fileToBase64Parts(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      const comma = dataUrl.indexOf(",");
      const base64 = comma >= 0 ? dataUrl.slice(comma + 1) : dataUrl;
      const mime = file.type || "image/jpeg";
      resolve({ base64, mime });
    };
    reader.onerror = () => reject(new Error("Could not read file."));
    reader.readAsDataURL(file);
  });
}

/**
 * Posts to Apps Script via HTML form into hidden iframe (avoids fetch/CORS).
 * Includes base64 image so the script can upload to Drive and store a clickable link.
 */
function submitToAppsScriptWebApp(payload) {
  if (
    !APPS_SCRIPT_WEB_APP_URL ||
    APPS_SCRIPT_WEB_APP_URL === "REPLACE_WITH_YOUR_APPS_SCRIPT_WEB_APP_URL"
  ) {
    throw new Error("Apps Script Web App URL is not set.");
  }

  const hiddenForm = document.createElement("form");
  hiddenForm.method = "POST";
  hiddenForm.action = APPS_SCRIPT_WEB_APP_URL;
  hiddenForm.target = "apps_script_iframe";
  hiddenForm.acceptCharset = "UTF-8";
  hiddenForm.style.display = "none";

  function addField(name, value) {
    const input = document.createElement("input");
    input.type = "hidden";
    input.name = name;
    input.value = value == null ? "" : String(value);
    hiddenForm.appendChild(input);
  }

  addField("fullName", payload.fullName);
  addField("phoneNumber", payload.phoneNumber);
  addField("email", payload.email);
  addField("age", String(payload.age));
  addField("gender", payload.gender);
  addField("heardHow", payload.heardHow);
  addField("heardPersonName", payload.heardPersonName || "");
  addField("paymentProofName", payload.paymentProofFile ? payload.paymentProofFile.name : "");
  addField("paymentProofBase64", payload.paymentProofBase64 || "");
  addField("paymentProofMime", payload.paymentProofMime || "");
  addField("termsAccepted", payload.termsAccepted ? "yes" : "");

  (payload.stationPrefs || []).forEach((station) => {
    addField("stationPrefs", station);
  });

  document.body.appendChild(hiddenForm);
  hiddenForm.submit();
  setTimeout(() => hiddenForm.remove(), 3000);
}

function initStationLimit() {
  const checkboxes = Array.from(
    document.querySelectorAll('input[type="checkbox"][data-station]')
  );

  const selectedStations = [];

  function updateCheckboxes() {
    const selectedCount = selectedStations.length;

    checkboxes.forEach((cb) => {
      const station = cb.getAttribute("data-station");
      const isSelected = selectedStations.includes(station);

      const shouldDisable = selectedCount >= 3 && !isSelected;
      cb.disabled = shouldDisable;
      const wrap = cb.closest(".station-option");
      if (wrap) wrap.setAttribute("aria-disabled", shouldDisable ? "true" : "false");
    });

    $("selectedCount").textContent = `${selectedStations.length}/3`;
    $("preferencesPreview").textContent =
      selectedStations.length ? selectedStations.join(" → ") : "None selected yet";
  }

  checkboxes.forEach((cb) => {
    cb.addEventListener("change", () => {
      const station = cb.getAttribute("data-station");
      const isNowChecked = cb.checked;

      if (isNowChecked) {
        if (selectedStations.length >= 3) {
          cb.checked = false;
          return;
        }
        selectedStations.push(station);
      } else {
        const idx = selectedStations.indexOf(station);
        if (idx !== -1) selectedStations.splice(idx, 1);
      }

      updateCheckboxes();
    });
  });

  updateCheckboxes();

  return {
    getSelectedStationsInOrder: () => selectedStations.slice(),
    reset: () => {
      selectedStations.length = 0;
      checkboxes.forEach((cb) => {
        cb.checked = false;
      });
      updateCheckboxes();
    },
  };
}

function validateForm({
  selectedStationsInOrder,
  paymentProofFile,
  heardHow,
  heardPersonName,
  termsAccepted,
}) {
  const errors = [];

  const fullName = trimValue($("fullName").value);
  const phoneNumber = trimValue($("phoneNumber").value);
  const email = trimValue($("email").value);
  const ageStr = trimValue($("age").value);
  const gender = trimValue($("gender").value);

  if (!fullName) errors.push("Please enter your full name.");
  if (!phoneNumber) errors.push("Please enter your phone number.");
  if (phoneNumber && phoneNumber.replace(/\D/g, "").length < 8)
    errors.push("Phone number looks too short. Please double-check.");
  if (!email) errors.push("Please enter your email address.");
  if (email && !isEmail(email)) errors.push("Please enter a valid email address.");
  const age = Number(ageStr);
  if (!ageStr || Number.isNaN(age)) errors.push("Please enter your age.");
  if (!Number.isNaN(age) && (age < 14 || age > 21))
    errors.push("Age must be between 14 and 21.");
  if (!gender) errors.push("Please select your gender.");

  if (!heardHow) errors.push("Please tell us how you heard about the event.");
  if (heardHow === HEARD_THROUGH_PERSON && !trimValue(heardPersonName)) {
    errors.push("Please enter the name of the person who told you about the event.");
  }

  if (!selectedStationsInOrder.length) errors.push("Please select at least 1 station.");
  if (selectedStationsInOrder.length > 3) errors.push("You can select a maximum of 3 stations.");

  if (!paymentProofFile) errors.push("Please upload your payment screenshot.");
  if (paymentProofFile && paymentProofFile.size > MAX_PAYMENT_IMAGE_BYTES) {
    errors.push(
      `Payment screenshot must be about ${Math.round(MAX_PAYMENT_IMAGE_BYTES / (1024 * 1024))}MB or less before upload.`
    );
  }

  if (!termsAccepted) errors.push("Please read and accept the event terms to continue.");

  return { errors, values: { fullName, phoneNumber, email, age, gender } };
}

function setSubmitState({ disabled }) {
  const btn = $("submitBtn");
  btn.disabled = disabled;
  if (disabled) btn.textContent = "Submitting…";
  else btn.textContent = "Submit Registration";
}

function toggleHeardPersonField() {
  const heardHow = $("heardHow").value;
  const wrap = $("heardPersonWrap");
  const input = $("heardPersonName");
  const isPerson = heardHow === HEARD_THROUGH_PERSON;
  wrap.hidden = !isPerson;
  input.required = isPerson;
  if (!isPerson) input.value = "";
}

document.addEventListener("DOMContentLoaded", () => {
  const stationState = initStationLimit();

  const form = $("registrationForm");
  const paymentProofInput = $("paymentProof");
  const heardHowSelect = $("heardHow");

  $("dismissSuccess")?.addEventListener("click", hideSuccess);
  $("dismissErrors")?.addEventListener("click", hideErrors);

  heardHowSelect.addEventListener("change", toggleHeardPersonField);
  toggleHeardPersonField();

  form.addEventListener("submit", async (e) => {
    e.preventDefault();

    hideErrors();
    hideSuccess();

    const selectedStationsInOrder = stationState.getSelectedStationsInOrder();
    const paymentProofFile =
      paymentProofInput.files && paymentProofInput.files[0] ? paymentProofInput.files[0] : null;
    const heardHow = heardHowSelect.value;
    const heardPersonName = trimValue($("heardPersonName").value);
    const termsAccepted = $("termsAgree").checked;

    const { errors, values } = validateForm({
      selectedStationsInOrder,
      paymentProofFile,
      heardHow,
      heardPersonName,
      termsAccepted,
    });

    if (errors.length) {
      showErrors(errors);
      const first = errors[0];
      if (first.includes("person who told")) $("heardPersonName").focus();
      else if (first.includes("full name")) $("fullName").focus();
      else if (first.includes("phone")) $("phoneNumber").focus();
      else if (first.includes("email")) $("email").focus();
      else if (first.includes("age")) $("age").focus();
      else if (first.includes("gender")) $("gender").focus();
      else if (first.includes("heard")) $("heardHow").focus();
      else if (first.includes("person")) $("heardPersonName").focus();
      else if (first.includes("station")) {
        document.querySelector("input[data-station]")?.focus?.();
      } else if (first.includes("screenshot")) paymentProofInput.focus();
      else if (first.includes("terms")) $("termsAgree").focus();
      return;
    }

    setSubmitState({ disabled: true });

    let uploadBlob = paymentProofFile;
    let uploadName = paymentProofFile.name;
    try {
      uploadBlob = await preparePaymentProofBlob(paymentProofFile);
      if (uploadBlob !== paymentProofFile) {
        uploadName = uploadName.replace(/\.[^.]+$/, "") + ".jpg";
      }
    } catch (prepErr) {
      console.error(prepErr);
      showErrors(["Could not process your payment image. Please use a JPG or PNG under 3 MB."]);
      setSubmitState({ disabled: false });
      return;
    }

    let proofBase64 = "";
    let proofMime = "";
    try {
      const parts = await fileToBase64Parts(uploadBlob);
      proofBase64 = parts.base64;
      proofMime = parts.mime || "image/jpeg";
    } catch (err) {
      console.error(err);
      showErrors(["Could not read your payment image. Please try another file."]);
      setSubmitState({ disabled: false });
      return;
    }

    const payload = {
      ...values,
      heardHow,
      heardPersonName: heardHow === HEARD_THROUGH_PERSON ? heardPersonName : "",
      stationPrefs: selectedStationsInOrder,
      paymentProofFile: { name: uploadName },
      paymentProofBase64: proofBase64,
      paymentProofMime: proofMime,
      termsAccepted: true,
    };

    try {
      submitToAppsScriptWebApp(payload);
      showSuccess(SUCCESS_MESSAGE);
      form.reset();
      stationState.reset();
      toggleHeardPersonField();
    } catch (err) {
      console.error(err);
      showErrors([
        "Submission failed. Check your Web App URL in app.js, redeploy Apps Script, and open the Registrations tab.",
      ]);
    } finally {
      setSubmitState({ disabled: false });
    }
  });
});
