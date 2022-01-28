window.addEventListener("load", load)

function load() {
    var inputs = document.getElementsByTagName("INPUT");

    backspacePolyfill()

    //Goes through each input
    for (var i = 0; i < inputs.length; i++) {
        var input = inputs[i]

        if (input.hasAttribute("placeholder"))
            addPlaceholderText(input)

        if (input.getAttribute("type") === "date") {
            input.addEventListener("keydown", function (event) { formatDate(event) })
            input.addEventListener("input", function (event) { ValidatePage1() })
            input.addEventListener("focusout", function (event) { onDateError(event.target, event) })
        }
    }
}

//Polyfill to listening to backspace events
function backspacePolyfill() {
    document.addEventListener('selectionchange', function (event) {
        var element = document.activeElement;

        if (element.tagName === 'INPUT' && element.type === 'text' && element.hasAttribute("placeholder")) {
            placeholderPolyfill(element.id)
        }
    });
}

//Polyfill for placeholder text on inputs
function placeholderPolyfill(id) {
    var input = document.getElementById(id)
    var placeholderText = document.getElementById(id + "-placeholder")

    if (!input || !placeholderText) return;

    if (input.hasAttribute("placeholder") && input.value != "")
        placeholderText.textContent = ""
    else
        placeholderText.textContent = input.getAttribute("placeholder")
}

//Attaches a placeholder text element over an input
function addPlaceholderText(input) {
    input.addEventListener("input", function (event) {
        placeholderPolyfill(event.target.id)
    })

    //Create placeholder text
    var placeHolderText = document.createElement("p");
    placeHolderText.textContent = input.getAttribute("placeholder");
    placeHolderText.className = "placeholder";
    placeHolderText.id = input.id + "-placeholder";

    var inputParent = input.parentElement;

    //Since elements can't be placed inside <input> tags, we wrap input in a div and add placeholder text to the div.
    var container = document.createElement("div");
    container.style.position = "relative";
    container.className = "form-field"
    container.style.padding = 0

    //Insert new wrapper div in position of input and place input inside container
    inputParent.insertBefore(container, input);
    inputParent.removeChild(input)

    container.appendChild(input);
    container.appendChild(placeHolderText);

    //IE9 doesnt allow you to disable collision on elements, so when we click the placeholder text, focus the input
    placeHolderText.addEventListener("click", function (event) {
        document.getElementById(input.id).select()
    })
}

function EffectiveDateFormatting() { }

function ValidatePage1() {
    var letterSelect = document.getElementById("selectLetter")
    var languageSelect = document.getElementById("selectLanguage")
    var dateInput = document.getElementById("inputEffectiveDate")

    if (letterSelect.value !== "--") {

        languageSelect.disabled = false

        if (selectLanguage.value !== "--") {

            dateInput.disabled = false

            if (dateInput.value.length === 10) {
                TogglePage1StartButton(dateValid(dateInput.value))
            }
            else {
                TogglePage1StartButton(false)
            }
        }
        else {
            dateInput.disabled = true
            TogglePage1StartButton(false)
        }
    }
    else {
        languageSelect.disabled = true
        dateInput.disabled = true
        TogglePage1StartButton(false)
    }
}

function formatDate(event) {

    var dateInput = event.target
    var value = event.target.value

    if (dateInput.value.length < 10 || event.key === "Backspace") {
        if (event.key === "Backspace") {
            if (value.length === 4 || value.length === 7) {
                dateInput.value = value.substring(0, value.length - 2);
                event.preventDefault()
            }

            TogglePage1StartButton(false)
        }
        else if (isNaN(event.key)) {
            event.preventDefault()
        }
        else {
            if (value.length === 2 || value.length === 5) {
                dateInput.value += "/"
            }
        }
    }
    else {
        event.preventDefault()
    }

    window.setTimeout(function () { onDateError(dateInput), 1 })
    window.setTimeout(ValidatePage1, 1)
    window.setTimeout(validateForm, 1)
}

function onDateError(dateInput, event) {
    if ((dateValid(dateInput.value) === false && dateInput.value.length === 10) || (event && event.type === "focusout" && dateInput.value.length < 10))
        showDateError(dateInput, true)
    else
        showDateError(dateInput, false)
}

function showDateError(dateInput, showError) {
    var containerElement = dateInput.parentElement.parentElement

    containerElement.className = containerElement.className.replace("input-error", " ")

    if (showError) {
        containerElement.className += " input-error"
    }
}

//Toggle Error
//All validation of forms in the letterbuilder
function validateForm() {
    var targetDate = document.getElementById("EffectiveDate")
    var targetElement = document.getElementById("DCPPlanNo")
    var targetCheckbox = document.getElementById("DCPStatus")
    var targetParent = targetElement.parentElement.parentElement

    var showGenerateButton = true

    targetParent.className = targetParent.className.replace("input-error", " ")

    if (dateValid(targetDate.value) === false)
        showGenerateButton = false

    if (targetElement.value === "") {
        if (targetCheckbox.checked)
            showGenerateButton = false

        targetParent.className += " input-error"
    }

    ToggleWordGenerateButtons(showGenerateButton)
}

function TogglePage1StartButton(isEnabled) {
    document.getElementById("page1ButtonStart").disabled = !isEnabled
}

//IE9 doesn't appear to have proper date validation (can have over 31 days in a month), so we include our own
function dateValid(date) {
    var year, month, day, split;

    split = date.split("/")

    year = parseInt(split[2]);
    month = parseInt(split[1]);
    day = parseInt(split[0]);

    if (!year || split[2].length < 4 || year > 9999 || year < 1) return false
    if (!month || month > 12 || month < 1) return false
    if (!day || day < 1) return false

    if (month % 2 === 1 && day > 31) return false
    if ((year % 4 !== 0 && month === 2 && day > 28) || (year % 4 === 0 && month === 2 && day > 29)) return false
    if (month % 2 === 0 && day > 30) return false

    return true
}
