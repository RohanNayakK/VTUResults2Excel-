const readFile = function(files) {
	const f = files[0];
	const reader = new FileReader();
	reader.onload = function(e) {
		let data = e.target.result;
		data = new Uint8Array(data);
    window.Bridge.sendReadExcel(data)

	};
	reader.readAsArrayBuffer(f);
};



const readIn = document.getElementById('readIn');
readIn.addEventListener('change', (e) => { readFile(e.target.files); }, false);