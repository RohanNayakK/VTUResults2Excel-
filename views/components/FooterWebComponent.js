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
background: #0e1419;
}

#projGroupName{
    color: goldenrod;
}
a{
    color: white;
}
</style>

<footer id="footer-section" >&copy All Rights Reserved - CEC - Designed & Developed by <b id="projGroupName">CEC/ISE/Project/2022-23/P04- I.S.E Dept</b>
<a href="../views/credits.html">->Credits (Click Here)</a>
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
