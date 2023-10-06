// noinspection DuplicatedCode

(function () {
    let tmpl = document.createElement('template');
    tmpl.innerHTML =
    `<button type="button" id="myBtn">Helper Button</button>` ;

    class CustomWidget extends HTMLElement {
        constructor() {
            super();
            this.init();
        }

        init() {

            let shadowRoot = this.attachShadow({mode: "open"});
            shadowRoot.appendChild(tmpl.content.cloneNode(true));
            this.addEventListener("click", event => {
            let newEvent = new Event("onClick");
            this.fireChanged();
            this.dispatchEvent(newEvent);
            });
        }

        fireChanged() {
            console.log("OnClick Triggered");
        }

    }

    customElements.define('ch-datart-test-ds-widget-1', CustomWidget);
})();