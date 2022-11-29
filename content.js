
let currentBrowser;
if (navigator.userAgent.indexOf("Chrome") !== -1) {
  currentBrowser = chrome;
} else {
  currentBrowser = browser;
}

document.addEventListener("loadMath", function (e) {
    const { isSpreadsheet } = e.detail;
    if (isSpreadsheet) {
      const scriptEl = document.createElement("script");
      const url = currentBrowser.runtime.getURL("spreadsheet_math.js");
      scriptEl.src = url;
      (document.head || document.documentElement).appendChild(scriptEl);
      scriptEl.onload = function () {
        scriptEl.parentNode.removeChild(scriptEl);
      };
    }
  });

const scriptEl3 = document.createElement("script");
const url3 = currentBrowser.runtime.getURL("load_math.js");
scriptEl3.src = url3;
(document.head || document.documentElement).appendChild(scriptEl3);
scriptEl3.onload = function () {
  scriptEl3.parentNode.removeChild(scriptEl3);
};
