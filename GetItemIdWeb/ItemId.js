(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            var itemElem = document.getElementById("itemId");
            var p = document.createElement("p");
            p.innerHTML = Office.context.mailbox.item.itemId;
            itemElem.appendChild(p);
        });
    };
})();