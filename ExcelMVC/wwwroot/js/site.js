// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

function updateFileText(input) {
    var fileText = input.files[0] ? input.files[0].name : "No file chosen";
    input.parentElement.querySelector('.file-text').textContent = fileText;
}