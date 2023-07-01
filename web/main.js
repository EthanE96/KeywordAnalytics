//Directory Input
 const dirPath = document.getElementById("dirInput");



//Onclick
document.querySelector("button").onclick = function () {
	//Call python funciton
	eel.synthesize_activate()(function (text) {
		//update div with message
		document.querySelector(".synthesize_activate").innerHTML = text;
	});

	//hides button afterwards
	document.querySelector("button").style.display = "none";
};