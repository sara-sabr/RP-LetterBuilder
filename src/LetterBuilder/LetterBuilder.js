window.addEventListener("load", load)

function load() {
    var inputs = document.getElementsByTagName("INPUT");

    backspacePolyfill()

    //Goes through each input
    for (var i = 0; i < inputs.length; i++) {
        if (inputs[i].hasAttribute("placeholder"))
            addPlaceholderText(inputs[i])
    }
}

//Polyfill to listening to backspace events
function backspacePolyfill() {
    document.addEventListener('selectionchange', function () {
        var element = document.activeElement;

        if (element.tagName === 'INPUT' && element.type === 'text')
            placeholderPolyfill(element.id)
    });
}

//Polyfill for placeholder text on inputs
function placeholderPolyfill(id) {
    var input = document.getElementById(id)
    var placeholderText = document.getElementById(id + "-placeholder")

    if (input.value != "")
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