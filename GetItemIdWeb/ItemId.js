(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            var itemElem = document.getElementById("itemId");
            var p = document.createElement("p");
            p.innerHTML = Office.context.mailbox.item.itemId;
            itemElem.appendChild(p);
            var br = document.createElement("br");
            itemElem.appendChild(br);
            var p2 = document.createElement("p");
            p2.innerHTML = "convertToRestId: <br>" + Office.context.mailbox.convertToEwsId(Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0);
            itemElem.appendChild(p2);
        });
    };
})();