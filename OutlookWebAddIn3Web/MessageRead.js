'use strict';



(function () {



    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            loadItemProps(Office.context.mailbox.item);
        });
    });




    var xmlhttp = new XMLHttpRequest();
    var url = "https://hospitalbedavailabilityapp.cfapps.us10.hana.ondemand.com/api/hospital/city";



    xmlhttp.onreadystatechange = function () {
        if (this.readyState == 4 && this.status == 200) {
            var myArr = JSON.parse(this.responseText);
            myFunction(myArr);
        }
    };



    xmlhttp.open("GET", url, true);
    xmlhttp.send();



    function myFunction(arr) {
        var out = [];
        var i;
        for (i = 0; i < arr.data.length; i++) {
            out.push(arr.data[i].state)
            /* console.log(arr.data[i]) */
        }
        let resultSet = [...new Set(out)];
        /* document.getElementById("id01").innerHTML = resultSet; */



        let stateSelect = document.getElementById("state");



        for (let i = 0; i < resultSet.length; i++) {
            //console.log(resultSet[i]);
            //stateSelect.options[stateSelect.options.length] = new Option(resultSet[i], resultSet[i]);
            stateSelect.options[stateSelect.options.length] = new Option(resultSet[i], resultSet[i]);
        }



    }



    function loadItemProps(item) {
        // Write message property values to the task pane
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-internetMessageId').text(item.internetMessageId);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");

    }
})();