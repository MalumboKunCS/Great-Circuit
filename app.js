// Apps Script Web App URL (Deploy → Web app → copy URL ending in /exec)
const APPS_SCRIPT_WEB_APP_URL =
  "https://script.google.com/macros/s/AKfycbwF43qtt9qneOpZblT5Jx89DI-gjpRLNgZ0WZM8_hn0YtMRNlceGDgPwiTU9k8-ia1dSA/exec";

function $(id) {
  return document.getElementById(id);
}

function trimValue(v) {
  return (v ?? "").toString().trim();
}

function isEmail(email) {
  // Simple, practical email check.
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

  box.style.display = "block";
  $("formSuccess").style.display = "none";
}

function showSuccess(msg) {
  const box = $("formSuccess");
  box.textContent = msg;
  box.style.display = box.classList.contains("toast") ? "inline-block" : "block";
  box.style.marginLeft = "auto";
  box.style.marginRight = "auto";
  $("formErrors").style.display = "none";
}

/**
 * Posts to Apps Script via a real HTML form into a hidden iframe.
 * fetch() from localhost or another site is often blocked by CORS/preflight, so doPost never runs.
 * This method behaves like a normal browser form submit and reliably reaches the web app.
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
    input.value = value;
    hiddenForm.appendChild(input);
  }

  addField("fullName", payload.fullName);
  addField("phoneNumber", payload.phoneNumber);
  addField("email", payload.email);
  addField("age", String(payload.age));
  addField("gender", payload.gender);
  addField(
    "paymentProofName",
    payload.paymentProofFile ? payload.paymentProofFile.name : ""
  );
  (payload.stationPrefs || []).forEach((station) => {
    addField("stationPrefs", station);
  });

  document.body.appendChild(hiddenForm);
  hiddenForm.submit();
  setTimeout(() => hiddenForm.remove(), 2000);
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
      cb.closest(".station-option").setAttribute("aria-disabled", shouldDisable ? "true" : "false");
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
          // Safety fallback 
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
      checkboxes.forEach((cb) => { cb.checked = false; });
      updateCheckboxes();
    },
  };
}

function validateForm({ selectedStationsInOrder, paymentProofFile }) {
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

  if (!selectedStationsInOrder.length) errors.push("Please select at least 1 station.");
  if (selectedStationsInOrder.length > 3) errors.push("You can select a maximum of 3 stations.");

  if (!paymentProofFile) errors.push("Please upload your payment screenshot.");
  if (paymentProofFile && paymentProofFile.size > 10 * 1024 * 1024)
    errors.push("Payment screenshot must be 10MB or less.");

  return { errors, values: { fullName, phoneNumber, email, age, gender } };
}

function setSubmitState({ disabled }) {
  const btn = $("submitBtn");
  btn.disabled = disabled;
  if (disabled) btn.textContent = "Submitting...";
  else btn.textContent = "Submit Registration";
}

document.addEventListener("DOMContentLoaded", () => {
  const stationState = initStationLimit();

  const form = $("registrationForm");
  const paymentProofInput = $("paymentProof");

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    showErrors([]); // hide if empty? we'll handle separately below

    $("formErrors").style.display = "none";
    $("formSuccess").style.display = "none";

    const selectedStationsInOrder = stationState.getSelectedStationsInOrder();
    const paymentProofFile = paymentProofInput.files && paymentProofInput.files[0] ? paymentProofInput.files[0] : null;

    const { errors, values } = validateForm({
      selectedStationsInOrder,
      paymentProofFile,
    });

    if (errors.length) {
      showErrors(errors);
      // Focus first invalid field for quick correction.
      const firstError = errors[0];
      if (firstError.includes("name")) $("fullName").focus();
      else if (firstError.includes("phone")) $("phoneNumber").focus();
      else if (firstError.includes("email")) $("email").focus();
      else if (firstError.includes("age")) $("age").focus();
      else if (firstError.includes("gender")) $("gender").focus();
      else if (firstError.includes("station")) {
        document.querySelector('input[data-station]:checked')?.focus?.();
      } else if (firstError.includes("screenshot")) paymentProofInput.focus();
      return;
    }

    setSubmitState({ disabled: true });

    const payload = {
      ...values,
      stationPrefs: selectedStationsInOrder,
      paymentProofFile,
    };

    try {
      submitToAppsScriptWebApp(payload);
      showSuccess("Thank you. Your registration has been received successfully and is now under verification. Your ticket will be sent to your email within 30 minutes.");
      form.reset();
      stationState.reset();
      setTimeout(() => { $("formSuccess").style.display = "none"; }, 6000);
    } catch (err) {
      console.error(err);
      showErrors([
        "Submission failed. Check your Web App URL in app.js, redeploy Apps Script, and look at the Registrations tab (not Form Responses).",
      ]);
    } finally {
      setSubmitState({ disabled: false });
    }
  });
});

