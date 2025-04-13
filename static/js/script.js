document.addEventListener("DOMContentLoaded", function() {
    // Tüm input, select ve buton elemanlarını aktif et
    let inputs = document.querySelectorAll('input, select, button');
    inputs.forEach(function(element) {
        element.disabled = false;
    });
});
