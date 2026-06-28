window.initSelect2 = (elementId, dotnetHelper) => {
    $(`#${elementId}`).select2({
        placeholder: "Pilih opsi...",
        allowClear: true
    }).on('change', function (e) {
        // Kirim balik data yang dipilih ke C#
        dotnetHelper.invokeMethodAsync('OnSelectChanged', $(this).val());
    });
};