fetch("data.json")
    .then(res => res.json())
    .then(data => {

        document.getElementById("total-pos").innerText =
            data.total_pos;

        document.getElementById("total-vendors").innerText =
            data.total_vendors;

        document.getElementById("total-containers").innerText =
            data.total_containers;

        new Chart(
            document.getElementById("importLocalChart"),
            {
                type: "doughnut",
                data: {
                    labels: data.import_local.map(
                        x => x["Import / Local"]
                    ),
                    datasets: [{
                        data: data.import_local.map(
                            x => x["PO#"]
                        )
                    }]
                }
            }
        );

        new Chart(
            document.getElementById("delayChart"),
            {
                type: "bar",
                data: {
                    labels: Object.keys(data.delay_counts),
                    datasets: [{
                        label: "Delay Count",
                        data: Object.values(data.delay_counts)
                    }]
                }
            }
        );

        new Chart(
            document.getElementById("vendorChart"),
            {
                type: "bar",
                data: {
                    labels: data.top_vendors.map(
                        x => x["Vendor Name"]
                    ),
                    datasets: [{
                        label: "PO Count",
                        data: data.top_vendors.map(
                            x => x["PO#"]
                        )
                    }]
                }
            }
        );

        new Chart(
            document.getElementById("countryChart"),
            {
                type: "bar",
                data: {
                    labels: data.countries.map(
                        x => x["Vendor Ctry"]
                    ),
                    datasets: [{
                        label: "PO Count",
                        data: data.countries.map(
                            x => x["PO#"]
                        )
                    }]
                }
            }
        );

    });