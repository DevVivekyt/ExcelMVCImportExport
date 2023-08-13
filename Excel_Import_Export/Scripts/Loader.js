
var startTime;
var timingElement = document.getElementById("timing");

function showLoader() {
    document.getElementById("loader").style.display = "flex";
    startTime = new Date().getTime();
    updateTiming();
}

function updateTiming() {
    var currentTime = new Date().getTime();
    var elapsedTime = currentTime - startTime;
    var seconds = Math.floor(elapsedTime / 1000);
    var minutes = Math.floor(seconds / 60);
    var hours = Math.floor(minutes / 60);

    timingElement.textContent = hours + "h " + (minutes % 60) + "m " + (seconds % 60) + "s";

    setTimeout(updateTiming, 1000); // Update timing every second
}

document.getElementById("uploadForm").addEventListener("submit", function () {
    showLoader();
});

// Hide the loader initially
window.onload = function () {
    document.getElementById("loader").style.display = "none";
};
