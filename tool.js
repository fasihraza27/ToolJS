let excelData = [];
let processedData = [];
let t0Businesses = new Map();
let processType = "t1";

document.getElementById("fileInput").addEventListener("change", function (e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, {
            type: "array",
             cellDates: true
        });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        excelData = XLSX.utils.sheet_to_json(worksheet);
        document.getElementById("result").innerHTML =
            `${excelData.length} records loaded successfully`;
    };
    reader.readAsArrayBuffer(file);
});

function calculateFee(amount) {
    amount = Number(amount);
    if (amount <= 5000) return 12.5;
    if (amount <= 10000) return 15.5;
    if (amount <= 100000) return 39;
    if (amount <= 250000) return 62.5;
    if (amount <= 1000000) return 125;
    if (amount <= 2500000) return 250;
    if (amount <= 5000000) return 375;
    return 500;
}

function pad(value) {
    return String(value).padStart(2, "0");
}

function formatDate(value) {
    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return "";
    }

    if (value instanceof Date) {
        if (isNaN(value.getTime())) {
            return "";
        }

        return (
            value.getFullYear() +
            "-" +
            pad(value.getMonth() + 1) +
            "-" +
            pad(value.getDate()) +
            " " +
            pad(value.getHours()) +
            ":" +
            pad(value.getMinutes()) +
            ":" +
            pad(value.getSeconds())
        );
    }

    value = String(value).trim();
    if (value.includes("T")) {
        value = value.replace("T", " ");
    }

    if (value.includes("/")) {
        const parts = value.split(/\s+/);
        const datePart = parts[0];
        let timePart = parts[1] || "00:00:00";
        let ampm = "";
        if (parts[2]) {
            ampm = parts[2].toUpperCase();
        }

        let [month, day, year] = datePart.split("/").map(Number);
        if (year < 100) {
            year += 2000;
        }

        let timeParts = timePart.split(":").map(Number);

        let hour = timeParts[0] || 0;
        let minute = timeParts[1] || 0;
        let second = timeParts[2] || 0;

        if (ampm === "PM" && hour < 12) {
            hour += 12;
        }

        if (ampm === "AM" && hour === 12) {
            hour = 0;
        }

        const date = new Date(
            year,
            month - 1,
            day,
            hour,
            minute,
            second
        );

        if (isNaN(date.getTime())) {
            return "";
        }

        return (
            date.getFullYear() +
            "-" +
            pad(date.getMonth() + 1) +
            "-" +
            pad(date.getDate()) +
            " " +
            pad(date.getHours()) +
            ":" +
            pad(date.getMinutes()) +
            ":" +
            pad(date.getSeconds())
        );
    }

    let match = value.match(
        /^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})$/
    );

    if (match) {
        const year = Number(match[1]);
        const month = Number(match[2]);
        const day = Number(match[3]);
        const hour = Number(match[4]);
        const minute = Number(match[5]);
        const second = Number(match[6]);

        const date = new Date(
            year,
            month - 1,
            day,
            hour,
            minute,
            second
        );

        return (
            date.getFullYear() +
            "-" +
            pad(date.getMonth() + 1) +
            "-" +
            pad(date.getDate()) +
            " " +
            pad(date.getHours()) +
            ":" +
            pad(date.getMinutes()) +
            ":" +
            pad(date.getSeconds())
        );
    }

    match = value.match(
        /^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})$/
    );

    if (match) {
        const year = Number(match[1]);
        const month = Number(match[2]);
        const day = Number(match[3]);
        const hour = Number(match[4]);
        const minute = Number(match[5]);

        const date = new Date(
            year,
            month - 1,
            day,
            hour,
            minute,
            0
        );

        return (
            date.getFullYear() +
            "-" +
            pad(date.getMonth() + 1) +
            "-" +
            pad(date.getDate()) +
            " " +
            pad(date.getHours()) +
            ":" +
            pad(date.getMinutes()) +
            ":00"
        );
    }

    if (!isNaN(value)) {
        const excelDate = Number(value);
        let date;

        if (
            typeof XLSX !== "undefined" &&
            XLSX.SSF &&
            XLSX.SSF.parse_date_code
        ) {

            const parsed = XLSX.SSF.parse_date_code(excelDate);
            if (parsed) {
                date = new Date(
                    parsed.y,
                    parsed.m - 1,
                    parsed.d,
                    parsed.H || 0,
                    parsed.M || 0,
                    Math.floor(parsed.S || 0)
                );
            }
        }

        if (!date) {
            const excelEpoch = new Date(1899, 11, 30);
            date = new Date(
                excelEpoch.getTime() +
                excelDate * 86400000
            );
        }

        if (isNaN(date.getTime())) {
            return "";
        }
        return (
            date.getFullYear() +
            "-" +
            pad(date.getMonth() + 1) +
            "-" +
            pad(date.getDate()) +
            " " +
            pad(date.getHours()) +
            ":" +
            pad(date.getMinutes()) +
            ":" +
            pad(date.getSeconds())
        );
    }
    return value;
}

function formatDateOnly(value) {
    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return "";
    }

    if (value instanceof Date) {
        if (isNaN(value.getTime())) {
            return "";
        }
        return (
            value.getFullYear() + "-" + pad(value.getMonth() + 1) + "-" + pad(value.getDate())
        );
    }
    value = String(value).trim();

    if (value.includes("/")) {
        const datePart = value.split(/\s+/)[0];
        let [month, day, year] = datePart.split("/").map(Number);
        if (year < 100) {
            year += 2000;
        }
        return (
            year +
            "-" +
            pad(month) +
            "-" +
            pad(day)
        );
    }

    if (!isNaN(value)) {
        const excelDate = Number(value);
        let date;
        if (
            typeof XLSX !== "undefined" &&
            XLSX.SSF &&
            XLSX.SSF.parse_date_code
        ) {
            const parsed = XLSX.SSF.parse_date_code(excelDate);
            if (parsed) {
                date = new Date(
                    parsed.y,
                    parsed.m - 1,
                    parsed.d
                );
            }
        }

        if (!date) {
            const excelEpoch = new Date(1899, 11, 30);
            date = new Date(
                excelEpoch.getTime() +
                excelDate * 86400000
            );
        }
        return (
            date.getFullYear() +
            "-" +
            pad(date.getMonth() + 1) +
            "-" +
            pad(date.getDate())
        );
    }
    return value;
}

function Transactiondatetime(value) {
    if(value === null || value === undefined || value === ""){
        return NaN;
    }

    if (value instanceof Date) {
        return value.getTime();
    }

    value = String(value).trim();
    if (value.includes("T")){
        value = value.replace("T", " ");
        if(value.split(":").length === 2){
            value += ":00";
        }
    }

    let match = value.match(
        /^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2}):(\d{2})$/
    )

    if(match){
        const year = Number(match[1]);
        const month = Number(match[2]);
        const day = Number(match[3]);
        const hour = Number(match[4]);
        const minute = Number(match[5]);
        const second = Number(match[6]);
        
        return new Date(year, month-1, day, hour, minute, second).getTime()
    }


    match = value.match(
        /^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})$/
    );


    if (match) {
        const year = Number(match[1]);
        const month = Number(match[2]);
        const day = Number(match[3]);
        const hour = Number(match[4]);
        const minute = Number(match[5]);
        return new Date(year, month - 1, day, hour, minute, 0).getTime();
    }

    if(value.includes("/")){
        const parts = value.split(/\s+/);
        const datePart = parts[0];
        let timePart = parts[1] || "00:00:00"
    }


    if (value.includes("/")) {

        const parts = value.split(/\s+/);
        const datePart = parts[0];
        let timePart = parts[1] || "00:00:00";
        let ampm = "";

        if (parts[2]) {
            ampm = parts[2].toUpperCase();
        }

        let [month, day, year] = datePart.split("/").map(Number);        
        if (year < 100) {
            year += 2000;
        }

        const timeParts = timePart.split(":").map(Number);
        let hour = timeParts[0] || 0;
        const minute = timeParts[1] || 0;
        const second = timeParts[2] || 0;

    if (ampm === "PM" && hour < 12 ) {
            hour += 12;
        }

        if ( ampm === "AM" && hour === 12) {
            hour = 0;
        }
        const date = new Date(
            year, month - 1, day, hour, minute, second
        );

        if (isNaN(date.getTime())) {
            return NaN;
        }
        return date.getTime();
    }

    if (!isNaN(value)) {
        const excelDate =
            Number(value);


        // Use SheetJS parser first

        if (
            typeof XLSX !== "undefined" &&
            XLSX.SSF &&
            XLSX.SSF.parse_date_code
        ) {
            const parsed =
                XLSX.SSF.parse_date_code(
                    excelDate
                );
            if (parsed) {
                const date =
                    new Date(
                        parsed.y,
                        parsed.m - 1,
                        parsed.d,
                        parsed.H || 0,
                        parsed.M || 0,
                        Math.floor(parsed.S || 0)
                    );
                return date.getTime();
            }
        }

        // Fallback
        const excelEpoch = new Date(1899, 11, 30);
        const date = new Date(
                excelEpoch.getTime() +
                excelDate * 86400000
            );
        return date.getTime();
    }
    return NaN;
}

function filterByTransactiondatetime(data) {
    const fromInput = document.getElementById("fromDateTime");
    const toInput = document.getElementById("toDateTime");

    const fromValue = fromInput.value.trim();
    const toValue = toInput.value.trim();

    if (!fromValue && !toValue) {
        return data;
    }

    let fromTime = -Infinity;
    let toTime = Infinity;

    if (fromValue) {
        fromTime = Transactiondatetime(fromValue);
        if (isNaN(fromTime)) {
            alert(
                "Invalid From date/time."
            );
            return null;
        }
    }

    if (toValue) {
        toTime = Transactiondatetime(toValue);
        if (isNaN(toTime)) {
            alert(
                "Invalid To date/time."
            );
            return null;
        }

        if (
            /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/
                .test(toValue)
        ) {

            toTime += (60 * 1000) - 1;
        }
    }

    if (fromTime > toTime) {
        alert(
            "From date/time cannot be greater than To date/time."
        );
        return null;
    }

    return data.filter((row) => {

        const value = row["Transaction Date Time"];
        
        if (
            value === null ||
            value === undefined ||
            value === ""
        ) {
            return false;
        }

        const transactionTime = Transactiondatetime(value);
        if (isNaN(transactionTime)) {
            console.log(
                "Invalid Transaction Date Time:",
                value
            );

            return false;
        }
        return (
            transactionTime >= fromTime &&
            transactionTime <= toTime
        );
    });
}

async function loadT0Businesses() {
    try {
        const response = await fetch(
            "https://api-ridercollection.foreebusiness.com/businessdata/get",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    settlement_period: "0"
                })
            }
        );

        const result = await response.json();
        if (!result.status) {
            alert(
                "Unable to fetch T+0 businesses."
            );
            return false;
        }
        t0Businesses.clear();
        result.data.forEach((item) => {
            const foreeId = String(item.business_id).slice(-4);
            t0Businesses.set(
                foreeId,
                item
            );
        });
        return true;
    } catch (error) {
        console.error("T+0 API Error:",error);
        alert("Unable to load T+0 businesses.");
        return false;
    }
}

function calculateFee0(amount, charges) {
    amount = Number(amount);
    const percentage =
        parseFloat(charges);
    if (isNaN(percentage)) {
        return 0;
    }
    return (
        amount *
        (percentage / 100)
    );
}

function processFile() {
    processType = "t1";
    if (excelData.length === 0) {
        alert(
            "Please upload a file first."
        );
        return;
    }

    const filteredData =
        filterByTransactiondatetime(
            excelData
        );

    if (filteredData === null) {
        return;
    }

    if (filteredData.length === 0) {
        alert(
            "No records found in the selected Transaction Date Time range."
        );
        return;
    }

    processedData = filteredData.map((row) => {

            let newRow = {...row};

            if (
                newRow["Bill Creation Date"]
            ) {

                newRow["Bill Creation Date"] = formatDateOnly(
                        newRow["Bill Creation Date"]
                    );
            }

            if (
                newRow["Bill Due Date"]
            ) {

                newRow["Bill Due Date"] =
                    formatDateOnly(newRow["Bill Due Date"]);
            }

            if (
                newRow["Transaction Date Time"]
            ) {

                newRow["Transaction Date Time"] =
                    formatDate(
                        newRow["Transaction Date Time"]
                    );
            }

            if (
                newRow["Initiator Settlement Date"]
            ) {
                newRow["Initiator Settlement Date"] =
                    formatDateOnly(
                        newRow["Initiator Settlement Date"]
                    );
            }

            const amount =
                Number( newRow["Applicable Amount"]) || 0;

            const fee = calculateFee(amount);
            newRow["Initiator Fee"] = fee;
            newRow["Tax On Initiator Fee"] = (fee * 0.15).toFixed(2);
            return newRow;
        });
    alert(
        `${processedData.length} T+1 rows processed successfully`
    );
}

async function processapi() {
    processType = "t0";
    if (excelData.length === 0) {
        alert(
            "Please upload a file first."
        );
        return;
    }

    const filteredData = filterByTransactiondatetime(
            excelData
        );

    if (filteredData === null) {
        return;
    }

    if (filteredData.length === 0) {
        alert(
            "No records found in the selected Transaction Date Time range."
        );
        return;
    }

    const loaded =
        await loadT0Businesses();

    if (!loaded) {
        return;
    }

    processedData = filteredData.filter((row) => {
                return t0Businesses.has(
                    String(row["Foree ID"])
                );
            })

            .map((row) => {
                let newRow = {
                    ...row
                };

                if (
                    newRow["Bill Creation Date"]
                ) {

                    newRow["Bill Creation Date"] = formatDateOnly(
                            newRow["Bill Creation Date"]
                        );
                }

                if (
                    newRow["Bill Due Date"]
                ) {

                    newRow["Bill Due Date"] = formatDateOnly(
                            newRow["Bill Due Date"]
                        );
                }

                if (
                    newRow["Transaction Date Time"]
                ) {
                    newRow["Transaction Date Time"] =
                        formatDate(
                            newRow["Transaction Date Time"]
                        );
                }

                if (
                    newRow["Initiator Settlement Date"]
                ) {
                    newRow["Initiator Settlement Date"] =
                        formatDateOnly(
                            newRow["Initiator Settlement Date"]
                        );
                }

                const business = t0Businesses.get(
                        String(newRow["Foree ID"])
                    );

                if (business) {
                    newRow["Business Name"] =
                        business.business_name;

                    const amount = Number(newRow["Applicable Amount"]) || 0;

                    const fee = calculateFee0(
                            amount,
                            business["1_bill_charges"]
                        );

                    newRow["Initiator Fee"] =
                        fee;

                    newRow["Tax On Initiator Fee"] =
                        fee * 0.15;
                }
                return newRow;
            });
    alert(
        `${processedData.length} T+0 rows processed successfully`
    );
}

function downloadCSV() {
    if (processedData.length === 0) {
        alert(
            "Process file first"
        );
        return;
    }

    const worksheet = XLSX.utils.json_to_sheet(
            processedData
        );

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "Processed Data"
    );

    let fileName;
    if (processType === "t0") {
        fileName =
            "processed_settlement_report_t0.csv";
    } else {
        fileName = "processed_settlement_report_t1.csv";
    }

    XLSX.writeFile(
        workbook,
        fileName
    );
}