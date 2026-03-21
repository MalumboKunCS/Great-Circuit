// Apps Script Web App URL (Deploy → Web app → copy URL ending in /exec)
const APPS_SCRIPT_WEB_APP_URL =
  "https://script.google.com/macros/s/AKfycbwF43qtt9qneOpZblT5Jx89DI-gjpRLNgZ0WZM8_hn0YtMRNlceGDgPwiTU9k8-ia1dSA/exec";

/** Max image size before base64 upload (Drive + Apps Script limits). ~3 MB raw file. */
const MAX_PAYMENT_IMAGE_BYTES = 3 * 1024 * 1024;

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
  $("formSuccess").classList.add("is-hidden");
}

function showSuccess(msg) {
  const box = $("formSuccess");
  box.textContent = msg;
  box.classList.remove("is-hidden");
  box.classList.add("toast--visible");
  $("formErrors").classList.add("is-hidden");
}

function fileToBase64Parts(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      const comma = dataUrl.indexOf(",");
      const base64 = comma >= 0 ? dataUrl.slice(comma + 1) : dataUrl;
      const mime = file.type || "image/png";
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
      `Payment screenshot must be ${Math.round(MAX_PAYMENT_IMAGE_BYTES / (1024 * 1024))}MB or less (try compressing the image).`
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

  heardHowSelect.addEventListener("change", toggleHeardPersonField);
  toggleHeardPersonField();

  form.addEventListener("submit", async (e) => {
    e.preventDefault();

    $("formErrors").classList.add("is-hidden");
    $("formSuccess").classList.add("is-hidden");
    $("formSuccess").classList.remove("toast--visible");

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

    let proofBase64 = "";
    let proofMime = "";
    try {
      const parts = await fileToBase64Parts(paymentProofFile);
      proofBase64 = parts.base64;
      proofMime = parts.mime;
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
      paymentProofFile,
      paymentProofBase64: proofBase64,
      paymentProofMime: proofMime,
      termsAccepted: true,
    };

    try {
      submitToAppsScriptWebApp(payload);
      showSuccess(
        "Thank you. Your registration has been received successfully and is now under verification. Your ticket will be sent to your email within 30 minutes."
      );
      form.reset();
      stationState.reset();
      toggleHeardPersonField();
      setTimeout(() => {
        $("formSuccess").classList.add("is-hidden");
        $("formSuccess").classList.remove("toast--visible");
      }, 6000);
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
