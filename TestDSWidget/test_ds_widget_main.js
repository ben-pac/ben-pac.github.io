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
            this._props = {};
        }
        onCustomWidgetAfterUpdate(changedProperties) {
            console.log('on custom widget after update', changedProperties)
            if ("testDataBinding" in changedProperties) {
                console.log(changedProperties)
                console.log(changedProperties.testDataBinding)
            }
        }
        onCustomWidgetBeforeUpdate(changedProperties) {
            console.log('on custom before', changedProperties)
            this._props = { ...this._props, ...changedProperties };
        }
        fireChanged() {
            console.log("OnClick Triggered");
            console.log('props', this._props)
        }

    }

    customElements.define('ch-datart-test-ds-widget-1', CustomWidget);
})();