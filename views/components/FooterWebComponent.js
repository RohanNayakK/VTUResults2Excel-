const footerTemplate = document.createElement('template');

footerTemplate.innerHTML = `
<style>
#footer-section {
text-align: center;
position: fixed;
bottom: 0;
left: 0;
width: 100vw;
height: 5vh;
}
</style>

<footer id="footer-section" >&copy All Rights Reserved - CEC - Designed & Developed by PO4(2023)- 
<a target="_blank" href="https://www.canaraengineering.in/is">ISE Department CEC</a>
<span id="footerLinks">
</span>
</footer>
`

class FooterWebComponent extends HTMLElement {
    constructor() {
        super();
        this._shadowRoot = this.attachShadow({ 'mode': 'open' });
        this._shadowRoot.appendChild(footerTemplate.content.cloneNode(true));

    }
}

window.customElements.define('app-footer', FooterWebComponent);
