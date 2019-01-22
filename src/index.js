'use strict';

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {

            $('#create').click(show_signature_pad);
            $('#save').click(insert);
        });
    };

    function show_signature_pad() {
        document.getElementById('signature-pad').removeAttribute("style");
    }

    function insert() {
        var dataURL = signaturePad.toDataURL("image/png");
        dataURL = dataURL.replace('data:image/png;base64,', '');
        Office.context.document.setSelectedDataAsync(dataURL, {
            coercionType: Office.CoercionType.Image,
            imageLeft: 100,
            imageTop: 100,
            imageWidth: 200,
            imageHeight: 200
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log("Action failed with error: " + asyncResult.error.message);

                } else {
                    hide_signature_pad();
                }
            });
    }
    function hide_signature_pad() {
        var canvas = document.getElementById('canvas');
        var ctx = canvas.getContext('2d');
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        document.getElementById('signature-pad').setAttribute("style", "visibility:hidden;");
    }
})();