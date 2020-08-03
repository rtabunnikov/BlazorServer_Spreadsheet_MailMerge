﻿function download(filename, contentType, bytesBase64) {
    var link = document.createElement('a');
    link.download = filename;
    link.href = "data:" + contentType + ";base64," + bytesBase64;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}