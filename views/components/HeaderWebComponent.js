const headerTemplate = document.createElement('template');

headerTemplate.innerHTML = `
<style>
#header-section {
display: flex;
position: fixed;
z-index: 2;
top: 0;
left: 0;
height: 20vh;
width: 100vw;
justify-content: center;
align-items: center;
gap: 20px;
background: #0e1419;
}

.twoNumTitle{
color: gold;
}

#logoImg{
height: 90%;
}

</style>


<header id="header-section">
  <img id="logoImg"  alt="CANARA" src="https://www.canaraengineering.in//assets/images/logo.png" />
  <h1>VTU Results<span class="twoNumTitle">2</span>Excel Tool</h1>
</header>
`

class HeaderWebComponent extends HTMLElement {
    constructor() {
        super();
        this._shadowRoot = this.attachShadow({ 'mode': 'open' });
        this._shadowRoot.appendChild(headerTemplate.content.cloneNode(true));

    }
}

window.customElements.define('app-header', HeaderWebComponent);
