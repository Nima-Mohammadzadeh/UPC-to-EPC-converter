document.getElementById('generate-serials').addEventListener('click', function() {
    const currentSerial = parseInt(document.getElementById('current-serial').textContent);
    const labelQuantity = parseInt(document.getElementById('label-quantity').value);

    if (!isNaN(labelQuantity) && labelQuantity > 0) {
        const startSerial = currentSerial + 1;
        const endSerial = currentSerial + labelQuantity;

        document.getElementById('start-serial').textContent = startSerial;
        document.getElementById('end-serial').textContent = endSerial;

        // Update the current serialization number at the top
        document.getElementById('current-serial').textContent = endSerial;
    } else {
        alert('Please enter a valid number of labels.');
    }
});
