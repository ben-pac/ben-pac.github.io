/*!

JSZip - A Javascript class for generating and reading zip files
<http://stuartk.com/jszip>

(c) 2009-2014 Stuart Knightley <stuart [at] stuartk.com>
Dual licenced under the MIT license or GPLv3. See https://raw.github.com/Stuk/jszip/master/LICENSE.markdown.

JSZip uses the library pako released under the MIT license :
https://github.com/nodeca/pako/blob/master/LICENSE
*/
var getScriptPromisify = (src) => {
  return new Promise((resolve) => {
    $.getScript(src, resolve);
  });
};

(function () {
  const template = document.createElement("template");
  template.innerHTML = `
    <style>
    :host {
    font-size: 13px;
    font-family: arial;
    overflow: auto;
    }
    </style>
    <section hidden>
    <article>
    <label for="fileUpload">Upload</label>
    
        <span></span><button id="remove">Remove</button>

    </article>
    <input hidden id="fileUpload" type="file" accept=".xls,.xlsx,.xlsm,.csv" />
    </section>
    `;

  class UploadXLS extends HTMLElement {
    constructor() {
      super();

      //HTML objects
      this.attachShadow({ mode: "open" });
      this.shadowRoot.appendChild(template.content.cloneNode(true));
      this._input = this.shadowRoot.querySelector("input");
      this._remove = this.shadowRoot.querySelector("#remove");

      //XLS related objects
      this._sheetNames = null; //holds array of Sheet Names
      this._data = null; //holds JSON Array returned from XLS sheet
      this._sheetName = "";
      this._measureNames = [];
      this._accountDimension = "";
      this._dateDimensions = [];
      this._dateValues = [];
      this._incomeAccts = [];
      this._reverseSignage = true;
      this._useFiscalDate = false;

      this._baseURL = window.location.origin;
      //EDIT THESE VALUES: Import Service values
      this._tokenURL = '';
      this._clientID = '';
      this._secret = '';
      this._useServiceAccount = false;

      //ODATA variables
      this._modelId=null;
      this._csrfToken = "";
      this._token = "";
      this._mapping = {};
      this._defaultValues = {};
      this._callPostData = false;
      this._rateTable = '';
      this._isCurrencyUpload = false;

      //Job variables
      this.failedRecords = [];
      this.status = "NOT_STARTED";
      this._currentBatch = 0;
      this._chunkSize = 100000;
      this._uploadResult = {
        currentStep: "Default",
        jobId: "",
        baseURL: "",
        modelId: "",
        mapping: {},
        defaultValues: {},
        data: [],
        httpStatus: "",
        totalNumberRowsInJob: 0,
        failedNumberRows: 0,
        failedRows: [],
        jobStatus: "",
        body: {},
        errorMessage: "",
      };
    }

    async uploadData(
      mapping,
      defaultValues,
      reverseSign,
      sheetName,
      useFiscalDate
    ) {
      // Check to see if the Base URL has been set and that a CSRF token is present
      if (this._baseURL.length === 0) {
        console.log("Your Base URL has not been set yet");
        return;
      }

      if (!this._csrfToken) {
        await this.getCSRFToken();
      }

      if (!this._csrfToken || this._csrfToken.length === 0) {
        console.log("There was an error retrieving your CSRF token");
        return;
      }
      if(!this._modelId){
        this._uploadResult.errorMessage = 'No Model has been set';
        this.dispatch("onFailedUpload");
        throw new Error('Model has not been set! Use the setModelId() function call to assign this value before calling the uploadData() function');
      }
      // Reset properties
      this.failedRecords = [];  
      this.status = "NOT_STARTED";

      this._reverseSignage = reverseSign;
      this._sheetName = sheetName;
      this._useFiscalDate = useFiscalDate;

      this._mapping = mapping;
      this._defaultValues = defaultValues;
      this._isCurrencyUpload = false;

      this._callPostData = true;
      this.showFileSelector();
    }
    async uploadCurrencyRates(
      rateTable,
      mapping,
      defaultValues,
      sheetName
    ) {
      // Check to see if the Base URL has been set and that a CSRF token is present
      if (this._baseURL.length === 0) {
        console.log("Your Base URL has not been set yet");
        return;
      }

      if (this._csrfToken.length < 1) {
        await this.getCSRFToken();
      }

      if (this._csrfToken.length === 0) {
        console.log("There was an error retrieving your CSRF token");
        return;
      }

      // Reset properties
      this.failedRecords = [];
      this.status = "NOT_STARTED";

      this._reverseSignage = false;
      this._sheetName = sheetName;
      this._useFiscalDate = false;

      this._mapping = mapping;
      this._defaultValues = defaultValues;
      this._measureNames=['rateValue'];
      if ('validFrom' in mapping) {
        this._dateDimensions=[mapping.validFrom];
    } else {
        this._dateDimensions=['validFrom'];
    }
      this._rateTable=rateTable;
      this._isCurrencyUpload = true;

      this._callPostData = true;
      this.showFileSelector();
    }

    async downloadFailedRecords() {
        const failedRows = this._uploadResult.failedRows;
        if (failedRows && failedRows.length > 0) {
            const outputData = [];
            outputData.push([...Object.keys(failedRows[0].row), "Reason"]);
            failedRows.forEach(d => {
                outputData.push([...Object.values(d.row), d.reason]);
            });
            const outputWorkbook = XLSX.utils.book_new();
            const outputWorksheet = XLSX.utils.aoa_to_sheet(outputData);
            XLSX.utils.book_append_sheet(outputWorkbook, outputWorksheet, "Failed Rows");
            XLSX.writeFile(outputWorkbook, "failed_records.xlsx");
        } else {
            console.log("No failed rows to download.");
        }
    }
    async downloadCustomFailedRecords(results) {

          if (!results || results.length === 0) {
            console.log("No records to download.");
            return;
        }

        // Extracting headers
        const headers = Object.keys(results[0])
            .filter(key => key !== "@MeasureDimension")
            .concat([results[0]['@MeasureDimension'].id]);
        
        const outputData = [headers]; // initial value with headers

        // Extracting rows
        results.forEach(result => {
            const row = headers.map(header => {
                if (header === results[0]['@MeasureDimension'].id) {
                    return result['@MeasureDimension'].rawValue;
                }
                return result[header].id;
            });
            outputData.push(row);
        });

        // Creating XLSX
        const outputWorkbook = XLSX.utils.book_new();
        const outputWorksheet = XLSX.utils.aoa_to_sheet(outputData);
        XLSX.utils.book_append_sheet(outputWorkbook, outputWorksheet, "Records");
        XLSX.writeFile(outputWorkbook, "records.xlsx");
    }
    showFileSelector() {
      this.handleRemove(); //remove any existing files, required if this action is run multiple times in the same session
      this._input.click();
    }

    validateBaseURL() {
      if (this._baseURL.endsWith("/")) {
        this._baseURL = this._baseURL.slice(0, -1);
      }
    }
    convertFiscalDate(fiscalDate) {
      if (!fiscalDate || (fiscalDate.length !== 6 && fiscalDate.length !== 7)) {
        throw new Error("Invalid fiscal date format.");
      }

      let desiredPeriod =
        fiscalDate.length === 7
          ? `${fiscalDate.substring(0, 4)}${fiscalDate.substring(5)}`
          : fiscalDate;

      let matchedDateObj = this._dateValues.find(
        (dateObj) => dateObj.FISCAL_CALPERIOD === desiredPeriod
      );

      if (!matchedDateObj) {
        throw new Error(
          `No matching date found for fiscal date: ${fiscalDate}`
        );
      }

      return matchedDateObj.CALMONTH;
    }

    async getMasterData(dimension) {
      const url =
        this._baseURL +
        "/api/v1/dataexport/providers/sac/" +
        this._modelId +
        `/${dimension}Master`;
        
        let options = {
          method: "GET",
          headers: {}
        };
      try{
        const data= await this.executeFetch(url,options);
        return data.value;
      }catch (error) {
        console.error('Fetch error:', error);
      }
      
      
    }

    async getAccountDimensions() {
      try {
        const url =
          this._baseURL +
          "/api/v1/dataexport/providers/sac/" +
          this._modelId +
          "/$metadata?$format=JSON";
      
        let options = {
          method: "GET",
          headers: {
          
          },
        };

        const data= await this.executeFetch(url,options);
        console.log(data);
        const factData = data["com.sap.cloudDataIntegration"]["FactData"];
      const masterData = data["com.sap.cloudDataIntegration"]["MasterData"];
      const edmNumberTypes = [
        "Edm.Decimal",
        "Edm.Integer",
        "Edm.Byte",
        "Edm.SByte",
        "Edm.Int16",
        "Edm.Int32",
        "Edm.Int64",
        "Edm.Single",
        "Edm.Double",
      ];

      for (let key in factData) {
        let value = factData[key];
        if (edmNumberTypes.includes(value["$Type"])) {
          this._measureNames.push(key);
        }
      }

      for (let key in masterData) {
        let value = masterData[key];
        if (value["@Integration.PropertyType"] === "ACCOUNT_TYPE") {
          let parentObject = key.split("___")[0]; // parsing the name before "___"
          this._accountDimension = parentObject;
        }
        if (value["$Type"] === "Edm.Date") {
          let parentObject = key.split("___")[0]; // parsing the name before "___"
          if (!this._dateDimensions.includes(parentObject)) {
            this._dateDimensions.push(parentObject);
            this._dateValues = await this.getMasterData(parentObject);
          }
        }
      }
      if (this._accountDimension) {
        // checking if this._accountDimension is populated
        const accountMasterData = await this.getMasterData(
          this._accountDimension
        );
        if (accountMasterData && Object.keys(accountMasterData).length > 0) {
          this._incomeAccts = accountMasterData
            .filter((obj) => obj.accType === "INC" || obj.accType === "LEQ")
            .map((obj) => obj.ID);
          console.log(this._incomeAccts);
        }
      }
      } catch (error) {
        console.error('Fetch error:', error);
      }    
    }

    async getCurrencyRateTableId() {
      const url =
        this._baseURL +
        "/api/v1/dataimport/currencyConversions";

        let options = {
          method: "GET",
          headers: {
            "x-csrf-token": this._csrfToken,
          },
        };

        let rateTableId='';
        try {
          const data= await this.executeFetch(url,options);
          for (let item of data.currencyConversions) {
            if (item.currencyConversionName === this._rateTable) {
              rateTableId = item.currencyConversionID;
              break;
            }
          }
          if (rateTableId) {
            console.log('Match found:', rateTableId);
          } else {
            console.log('No match found');
          }
        } catch (error) {
          console.error('Error fetching the data:', error);
          throw new Error('Error fetching the data:', error);
        }
      return rateTableId;
    }

    async getCSRFToken() {
      // First, fetch the CSRF token
      let url = this._baseURL + "/api/v1/csrf";
      //let url = this._baseURL +"/api/v1/dataimport/jobs";
      let csrfOptions = {
        method: "GET",
        headers: {
          "x-csrf-token": "Fetch",
        },
      };
      try {
        let csrfResponse = await fetch(url, csrfOptions);
        if (!csrfResponse.ok) {
          throw new Error(`HTTP error! status: ${csrfResponse.status}`);
        }
        // Log all headers
        for (let [key, value] of csrfResponse.headers.entries()) {
          console.log(`${key}: ${value}`);
        }
        this._csrfToken = csrfResponse.headers.get("x-csrf-token");
      } catch (error) {
        console.error(
          "There was a problem with the fetch operation: " + error.message
        );
      }
      return true;
    }
    async getAccessToken() {
      const url = this._tokenURL+'/?grant_type=client_credentials&x-sap-sac-custom-auth=true';
    //'Authorization': 'Basic ' + btoa(`${id}:${secret}`)
    //'Authorization': 'Basic ' + btoa('sb-43eae37d-563b-4c59-87b8-98655a271e31!b164730|client!b655:3d2e816d-42ae-4df7-be03-aa1c0ddd9d80$NNaaD7qv5GkKruaoM6ixrjwmIe0C8MF5P5v1OMcSJxE=')
    let options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Authorization': 'Basic ' + btoa(`${this._clientID}:${this._secret}`)
      },
    };
    let response = await fetch(url, options);
    let json = await response.json();
    this._token= json.access_token;
    return json.access_token;     
    }
    async populateAccessTokens() {
      await this.getAccessToken
      await this.getCSRFToken();
      return this._csrfToken;
    }

    
    setNames(sheetNames) {
      this._sheetNames = sheetNames;
    }
    setBaseURL(url){
      this._baseURL = url;
    }

    setData(newData) {
      if (newData !== undefined) {
        this._data = newData;
        console.log(newData);
        this.dispatch("onFileUpload");
      }
    }
    resetData() {
      this._callPostData = false;
      this._sheetNames = null; //holds array of Sheet Names
      this._data = null; //holds JSON Array returned from XLS sheet
      this._sheetName = "";
    }
    parseCSV(csvData) {
      var lines = csvData.split("\n");
      var headers = lines[0].split(","); // assuming comma-separated CSV
      var sheetData = [];

      for (let i = 1; i < lines.length; i++) {
        // skip the headers line
        if (lines[i].trim() === "") continue; // skip empty lines

        let obj = {};
        let row = lines[i].split(",");
        var reverse = false; // Initialize reverse for each row

        for (let j = 0; j < headers.length; j++) {
          let key = headers[j].trim();
          let value = row[j] ? row[j].toString().trim() : ""; // Convert value to string and then trim

          if (this._reverseSignage) {
            if (
              key === this._accountDimension ||
              key === this._mapping[this._accountDimension]
            ) {
              if (
                this._incomeAccts.includes(value) ||
                this._defaultValues[this._accountDimension] === value
              ) {
                reverse = true;
              }
            }
          }

          if (
            (this._measureNames.includes(key) ||
              (Object.values(this._mapping).includes(key) &&
                this._measureNames.includes(
                  Object.keys(this._mapping).find((k) => this._mapping[k] === key)
                ))) &&
            !isNaN(value)
          ) {
            if (reverse) {
              obj[key] = Number((Number(value) * -1).toFixed(7));
            } else {
              obj[key] = Number(Number(value).toFixed(7));
            }
          } else {
            //check to see if it as Date dimension. If so ensure that it is passed through in the correct format
            if (this._dateDimensions.includes(key)) {
              let date;
              if (value.includes("/")) {
                // M/D/YYYY format
                let dateComponents = value.split("/");
                if (dateComponents.length === 3) {
                  date = new Date(
                    dateComponents[2],
                    dateComponents[0] - 1,
                    dateComponents[1]
                  );
                }
              /*} else if (value.length === 6 && !this._useFiscalDate) {
                // YYYYMM format
                let year = value.slice(0, 4);
                let month = value.slice(4) - 1; // JavaScript months are 0-indexed
                date = new Date(year, month);*/
              } else if (this._useFiscalDate) {
                //apply Fiscal Year settings
                value = this.convertFiscalDate(value);
              }
              if (date) {
                value = date.toISOString().split("T")[0]; // format as YYYY-MM-DD
              }
            }

            obj[key] = String(value);
          }
        }
        sheetData.push(obj);
      }

      this.setNames(["Sheet1"]);
      this.setData({ Sheet1: sheetData }); // mimic the existing behavior for XLSX files
      this.handleRemove();
    }

    async parseExcel(file) {
      const temp = this;
      var reader = new FileReader();
      var fileType = file.name.substring(file.name.lastIndexOf(".") + 1);

      reader.onload = function (e) {
        var data = e.target.result;
        var workbook;

        if (fileType === "csv") {
          var arr = new TextDecoder("utf-8").decode(new Uint8Array(data));
          temp.parseCSV(arr);
        } else {
          workbook = XLSX.read(data, { type: "binary", cellDates: true });
          var sheetData = [];
          var sheetNames = [];
          workbook.SheetNames.forEach(function (sheetName) {
            var XL_row_object = XLSX.utils.sheet_to_row_object_array(
              workbook.Sheets[sheetName]
            );
            XL_row_object.forEach(function (row) {
              var reverse = false;
              for (var key in row) {
                var value = row[key] ? row[key] : ""; 
                if (value instanceof Date) {
                    // Handle Date object (keep it as a Date object for now)
                    row[key] = value.toISOString().substring(0, 10);
                } else {
                    // If it's not a Date object, convert it to a string and trim it
                    value = value.toString().trim();
                }

                if (temp._reverseSignage) {
                  if (
                    key === temp._accountDimension ||
                    key === temp._mapping[temp._accountDimension]
                  ) {
                    if (
                      temp._incomeAccts.includes(row[key]) ||
                      temp._defaultValues[temp._accountDimension] === row[key]
                    ) {
                      reverse = true;
                    }
                  }
                }
                if (
                  (temp._measureNames.includes(key) ||
                    (Object.values(temp._mapping).includes(key) &&
                      temp._measureNames.includes(
                        Object.keys(temp._mapping).find((k) => temp._mapping[k] === key)
                      ))) &&
                  !isNaN(row[key])
                ) {
                  if (reverse) {
                    row[key] = Number((Number(row[key]) * -1).toFixed(7));
                  } else {
                    row[key] = Number(Number(row[key]).toFixed(7));
                  }
                } else {
                  // If value is a Date object, format it
                  if (temp._dateDimensions.includes(key)) {
                    if (value instanceof Date) {
                      row[key] = value.toISOString().substring(0, 10);
                    } else {
                      let date;
                      if (value.includes("/")) {
                        // M/D/YYYY format
                        let dateComponents = value.split("/");
                        if (dateComponents.length === 3) {
                          date = new Date(
                            dateComponents[2],
                            dateComponents[0] - 1,
                            dateComponents[1]
                          );
                        }
                      /*} 
                      else if (value.length === 6 && !temp._useFiscalDate) {
                        // YYYYMM format
                        let year = value.slice(0, 4);
                        let month = value.slice(4) - 1; // JavaScript months are 0-indexed
                        date = new Date(year, month);
                        */
                      } else if (temp._useFiscalDate) {
                        //apply Fiscal Year settings
                        row[key] = temp.convertFiscalDate(value);
                      }
                      if (date) {
                        row[key] = date.toISOString().split("T")[0]; // format as YYYY-MM-DD
                      }
                    }
                  } else {
                    row[key] = String(row[key]);
                  }
                }
              }
            });
            var json_object = JSON.stringify(XL_row_object);
            var rowData = JSON.parse(json_object);
            sheetNames.push(sheetName);
            sheetData[sheetName] = rowData;
          });
          temp.setNames(sheetNames);
          temp.setData(sheetData);
          temp.handleRemove();
        }

        reader = null;
      };

      reader.onerror = function (ex) {
        reader = null;
        console.log(ex);
      };

      if (fileType === "csv") {
        reader.readAsArrayBuffer(file);
      } else {
        reader.readAsBinaryString(file);
      }
    }

    async createJobAndPostDataInChunks(data) {
      let jobId;
      let jobURL;
      let runJobURL;
      let validateJobURL;
      let invalidRowsURL;
      let blnRunJob = true;
      let totalNumberRowsInJob;
      let failedNumberRows;
      let status;
      this._uploadResult = {
        currentStep: "Initialization",
        jobId: "",
        baseURL: "",
        modelId: this._modelId,
        mapping: this._mapping,
        defaultValues: this._defaultValues,
        data: data,
        httpStatus: "",
        totalNumberRowsInJob: data.length,
        failedNumberRows: 0,
        failedRows: [],
        jobStatus: "",
        body: {},
        errorMessage: "",
      };

      // Create the job
      try {
        this._uploadResult.currentStep = "Creating Job";
        let jobCreationResponse = await this.createJob();
        jobId = jobCreationResponse.jobID;
        jobURL = jobCreationResponse.jobURL;
        this._uploadResult.jobId = jobId;
      } catch (error) {
        this._uploadResult.jobStatus = "FAILED";

        console.error(
          `Error creating job: ${error.message}\n ${this._uploadResult.body.error.message}`
        );
        this.dispatch("onFailedUpload");
        return this._uploadResult;
      }

      // Split data and post in chunks
      this._uploadResult.currentStep = "Posting Data to Job";
      for (let i = 0; i < data.length; i += this._chunkSize) {
        let chunk = data.slice(i, i + this._chunkSize);
        try {
          let postDataResponse = await this.postDataToJob(jobURL, chunk);
          this._uploadResult.failedNumberRows +=
            postDataResponse.failedNumberRows;
          this._uploadResult.failedRows = this._uploadResult.failedRows.concat(
            this.addRowAsString(postDataResponse.failedRows)
          );

          if (postDataResponse.upsertedNumberRows === 0) {
            this._uploadResult.jobStatus = "FAILED";
            this.dispatch("onFailedUpload");
            return this._uploadResult;
          }
          this._currentBatch = i;
          this.dispatch("onBatchUpload");
          runJobURL = postDataResponse.runJobURL;
          validateJobURL = postDataResponse.validateJobURL;
        } catch (error) {
          console.error(
            `Error posting data: ${error.message}\n ${this._uploadResult.body.error.message}`
          );
          this._uploadResult.jobStatus = "FAILED";
          this._uploadResult.httpStatus = error.message;
          this.dispatch("onFailedUpload");
          return this._uploadResult;
        }
      }

      //validate the job
      try {
        this._uploadResult.currentStep = "Validating Job";
        let validateJobResponse = await this.validateJob(validateJobURL);
        invalidRowsURL = validateJobResponse.invalidRowsURL;
        totalNumberRowsInJob = validateJobResponse.totalNumberRowsInJob;
        failedNumberRows = validateJobResponse.failedNumberRows;
        this._uploadResult.failedNumberRows =
          this._uploadResult.failedNumberRows + failedNumberRows;

        if (failedNumberRows === totalNumberRowsInJob) {
          //don't run Job as it will error out
          blnRunJob = false;
          status = "FAILED";
          this._uploadResult.jobStatus = status;
        }

        if (failedNumberRows > 0) {
          let failedRecordsResponse = await this.getFailedRecords(
            invalidRowsURL
          );
          let transformedResponse = this.transformResponse(
            failedRecordsResponse.failedRows
          );

          this._uploadResult.failedRows =
            this._uploadResult.failedRows.concat(transformedResponse);
        }
      } catch (error) {
        this._uploadResult.jobStatus = "FAILED";
        this._uploadResult.httpStatus = error.message;
        console.error(
          `Error posting data: ${error.message}\n ${this._uploadResult.body.error.message}`
        );
        this.dispatch("onFailedUpload");
        return this._uploadResult;
      }

      //run the Job
      if (blnRunJob) {
        try {
          this._uploadResult.currentStep = "Running Job";
          await this.runJob(runJobURL);
        } catch (error) {
            this._uploadResult.jobStatus = "FAILED";
            this._uploadResult.httpStatus = error.message;
          console.error(
            `Error running job: ${error.message}\n ${this._uploadResult.body.error.message}`
          );
            this.dispatch("onFailedUpload");
            return this._uploadResult;
        }

        // Check job status
        do {
          try {
            this._uploadResult.currentStep = "Checking Job Status";
            status = await this.checkJobStatus(jobURL);
          } catch (error) {
            this._uploadResult.jobStatus = "FAILED";
            this._uploadResult.httpStatus = error.message;
            console.error(`Error checking job status: ${error.message}\n ${this._uploadResult.body.error.message}`);
            this.dispatch("onFailedUpload");
            return this._uploadResult;
          }
          
          if (status === "COMPLETED") {
            if (failedNumberRows > 0) {
              status = "COMPLETED_WITH_FAILURES";
            }
            this.dispatch("onDataUpload");
          } else if (status === "FAILED") {
            this.dispatch("onFailedUpload");
          }
        } while (
          status !== "COMPLETED" &&
          status !== "FAILED" &&
          status !== "COMPLETED_WITH_FAILURES"
        );
        this._uploadResult.jobStatus=status;
        this._uploadResult.currentStep="Upload Completed";
      
      return this._uploadResult;
      }else{
        this.dispatch("onFailedUpload");
        return this._uploadResult;
      }
    }
    async createJobAndPostCurrencyData(data) {
      let jobId;
      let jobURL;
      let runJobURL;
      let validateJobURL;
      let invalidRowsURL;
      let blnRunJob = true;
      let totalNumberRowsInJob;
      let failedNumberRows;
      let status;
      let rateTableId='';
      this._uploadResult = {
        currentStep: "Initialization",
        jobId: "",
        baseURL: "",
        modelId: this._rateTable,
        mapping: this._mapping,
        defaultValues: this._defaultValues,
        data: data,
        httpStatus: "",
        totalNumberRowsInJob: data.length,
        failedNumberRows: 0,
        failedRows: [],
        jobStatus: "",
        body: {},
        errorMessage: "",
      };

      //Get Rate Table ID
      try{
        this._uploadResult.currentStep = "Getting Rate Table ID";
        rateTableId = await this.getCurrencyRateTableId();
        if(rateTableId.length===0){
          this._uploadResult.jobStatus = "FAILED";
          this._uploadResult.errorMessage = "Could not find Currency Rate Table with Name: "+ this._rateTable;
          this.dispatch("onFailedUpload");
          return this._uploadResult;
        }

      }catch (error) {
        this._uploadResult.jobStatus = "FAILED";
        console.error(
          `Error creating job: ${error.message}\n ${this._uploadResult.body.error.message}`
        );
        this.dispatch("onFailedUpload");
        return this._uploadResult;
      }

      // Create the job
      try {
        this._uploadResult.currentStep = "Creating Job";
        let jobCreationResponse = await this.createCurrencyJob(rateTableId);
        jobId = jobCreationResponse.jobID;
        jobURL = jobCreationResponse.jobURL;
        this._uploadResult.jobId = jobId;
      } catch (error) {
        this._uploadResult.jobStatus = "FAILED";

        console.error(
          `Error creating job: ${error.message}\n ${this._uploadResult.body.error.message}`
        );
        this.dispatch("onFailedUpload");
        return this._uploadResult;
      }

      // Split data and post in chunks
      this._uploadResult.currentStep = "Posting Data to Job";
      for (let i = 0; i < data.length; i += this._chunkSize) {
        let chunk = data.slice(i, i + this._chunkSize);
        try {
          let postDataResponse = await this.postDataToJob(jobURL, chunk);
          this._uploadResult.failedNumberRows +=
            postDataResponse.failedNumberRows;
          this._uploadResult.failedRows = this._uploadResult.failedRows.concat(
            this.addRowAsString(postDataResponse.failedRows)
          );

          if (postDataResponse.upsertedNumberRows === 0) {
            this._uploadResult.jobStatus = "FAILED";
            this.dispatch("onFailedUpload");
            return this._uploadResult;
          }
          this._currentBatch = i;
          this.dispatch("onBatchUpload");
          runJobURL = postDataResponse.runJobURL;
          validateJobURL = postDataResponse.validateJobURL;
        } catch (error) {
          console.error(
            `Error posting data: ${error.message}\n ${this._uploadResult.body.error.message}`
          );
          this._uploadResult.jobStatus = "FAILED";
          this._uploadResult.httpStatus = error.message;
          this.dispatch("onFailedUpload");
          return this._uploadResult;
        }
      }

      //validate the job
      try {
        this._uploadResult.currentStep = "Validating Job";
        let validateJobResponse = await this.validateJob(validateJobURL);
        invalidRowsURL = validateJobResponse.invalidRowsURL;
        totalNumberRowsInJob = validateJobResponse.totalNumberRowsInJob;
        failedNumberRows = validateJobResponse.failedNumberRows;
        this._uploadResult.failedNumberRows =
          this._uploadResult.failedNumberRows + failedNumberRows;

        if (failedNumberRows === totalNumberRowsInJob) {
          //don't run Job as it will error out
          blnRunJob = false;
          status = "FAILED";
          this._uploadResult.jobStatus = status;
        }

        if (failedNumberRows > 0) {
          let failedRecordsResponse = await this.getFailedRecords(
            invalidRowsURL
          );
          let transformedResponse = this.transformResponse(
            failedRecordsResponse.failedRows
          );

          this._uploadResult.failedRows =
            this._uploadResult.failedRows.concat(transformedResponse);
        }
      } catch (error) {
        this._uploadResult.jobStatus = "FAILED";
        this._uploadResult.httpStatus = error.message;
        console.error(
          `Error posting data: ${error.message}\n ${this._uploadResult.body.error.message}`
        );
        this.dispatch("onFailedUpload");
        return this._uploadResult;
      }

      //run the Job
      if (blnRunJob) {
        try {
          this._uploadResult.currentStep = "Running Job";
          await this.runJob(runJobURL);
        } catch (error) {
            this._uploadResult.jobStatus = "FAILED";
            this._uploadResult.httpStatus = error.message;
          console.error(
            `Error running job: ${error.message}\n ${this._uploadResult.body.error.message}`
          );
            this.dispatch("onFailedUpload");
            return this._uploadResult;
        }

        // Check job status
        do {
          try {
            this._uploadResult.currentStep = "Checking Job Status";
            status = await this.checkJobStatus(jobURL);
          } catch (error) {
            this._uploadResult.jobStatus = "FAILED";
            this._uploadResult.httpStatus = error.message;
            console.error(`Error checking job status: ${error.message}\n ${this._uploadResult.body.error.message}`);
            this.dispatch("onFailedUpload");
            return this._uploadResult;
          }
          
          if (status === "COMPLETED") {
            if (failedNumberRows > 0) {
              status = "COMPLETED_WITH_FAILURES";
            }
            this.dispatch("onDataUpload");
          } else if (status === "FAILED") {
            this.dispatch("onFailedUpload");
          }
        } while (
          status !== "COMPLETED" &&
          status !== "FAILED" &&
          status !== "COMPLETED_WITH_FAILURES"
        );
        this._uploadResult.jobStatus=status;
        this._uploadResult.currentStep="Upload Completed";
      
      return this._uploadResult;
      }else{
        this.dispatch("onFailedUpload");
        return this._uploadResult;
      }
    }
    async createJob() {
      let url =
        this._baseURL +
        "/api/v1/dataimport/models/" +
        this._modelId +
        "/factData";
      let options = {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-csrf-token": this._csrfToken
        },
        body: JSON.stringify({
          Mapping: this._mapping,
          DefaultValues: this._defaultValues,
        }),
      };

      return await this.executeFetch(url, options);
    }
    async createCurrencyJob(rateTableId) {
      let url =
        this._baseURL +
        "/api/v1/dataimport/currencyConversions/" +
        rateTableId;
      let options = {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-csrf-token": this._csrfToken
        },
        body: JSON.stringify({
          Mapping: this._mapping,
          DefaultValues: this._defaultValues,
        }),
      };

      return await this.executeFetch(url, options);
    }

    async postDataToJob(url, data) {
      let options = {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-csrf-token": this._csrfToken
        },
        body: JSON.stringify({
          Mapping: this._mapping,
          DefaultValues: this._defaultValues,
          Data: data,
        }),
      };

      return await this.executeFetch(url, options);
    }
    async validateJob(validateJobURL) {
        let options = {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-csrf-token': this._csrfToken
            }
        };

        return await this.executeFetch(validateJobURL, options);

    }
    async getFailedRecords(invalidRowsURL) {
      let options = {
        method: "GET",
        headers: { "Content-Type": "application/json"},
      };

      return await this.executeFetch(invalidRowsURL, options);
    }
    async runJob(runJobURL) {
      let options = {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-csrf-token": this._csrfToken
        },
      };

      return await this.executeFetch(runJobURL, options);
    }

    async checkJobStatus(jobURL) {
      let url = jobURL + "/status";
      let options = {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
          "x-csrf-token": this._csrfToken
        },
      };

      let json = await this.executeFetch(url, options);
      return json.jobStatus;
    }

    async executeFetch(url, options) {
      // If this._useServiceAccount is true, add the extra headers
      if (this._useServiceAccount) {
        options.headers = {
          ...options.headers, // Spread existing headers if any
          'Authorization': 'Bearer ' + this._token,
          'x-sap-sac-custom-auth': true
        };
      }
      let response = await fetch(url, options);
      if (!response.ok) {
        this._uploadResult.errorMessage = this._uploadResult.body.error.message;
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      let json = await response.json();
      this._uploadResult.body = json;
      this._uploadResult.httpStatus = response.status;
    
      // Get the x-correlationid from the response headers and log it to the console
      let correlationID = response.headers.get('x-correlationid');
      if (correlationID) {
        console.log('Current Step:', this._uploadResult.currentStep, ' - Correlation ID: ', correlationID);
      }
      return json;
    }
    

    addRowAsString(array) {
      return array.map((item) => {
        // Convert the row object to a string
        let rowAsString = this.stringifyJSON(item.row);

        // Return a new object that includes the rowAsString property
        return {
          ...item,
          rowAsString: rowAsString,
        };
      });
    }
    transformResponse(failedRows) {
      return failedRows.map((item) => {
        // Create a copy of the item
        let newItem = { ...item };

        // Remove _REJECTION_REASON from the new item and assign its value to reason
        let reason = newItem._REJECTION_REASON;
        delete newItem._REJECTION_REASON;

        // Convert row to string
        let rowAsString = this.stringifyJSON(newItem);

        // Return the transformed item
        return {
          row: newItem,
          rowAsString: rowAsString,
          reason: reason,
        };
      });
    }
    stringifyJSON(json){
        let result = '';
        for (let [key, value] of Object.entries(json)) {
            result += `${key}: ${value}\n`;
        }
        return result;
      }
    //events

    //triggered when a user removes the Excel file
    handleRemove() {
      const el = this._input;
      const file = el.files[0];
      el.value = "";
      this.dispatch("change", file);
    }
    async handleFileSelect(evt) {
      console.log(Date.now()); //prints timestamp to console...for testing purposes only
      var files = evt.target.files; // FileList object

      this.setData(await this.parseExcel(files[0]));
    }

    dispatch(event, arg) {
      //this.dispatchEvent(new CustomEvent(event, {detail: arg}));

      if (event === "onFileUpload" && this._callPostData) {
        if (this._sheetName.length < 1) {
          this._sheetName = this._sheetNames[0];
        }
        const data = this._data[this._sheetName];
        if(this._isCurrencyUpload){
          this.createJobAndPostCurrencyData(data);
        }else{
          this.createJobAndPostDataInChunks(data);
        }        
        this.dispatchEvent(new CustomEvent(event, { detail: arg }));
      } else {
        this.dispatchEvent(new CustomEvent(event, { detail: arg }));
      }
    }
    largeUint8ArrayToString(u8a) {
      var CHUNK_SZ = 0x8000;
      var c = [];
      for (var i = 0; i < u8a.length; i += CHUNK_SZ) {
        c.push(String.fromCharCode.apply(null, u8a.subarray(i, i + CHUNK_SZ)));
      }
      return c.join("");
    }

    //setters and getters
    async setModelId(modelId) {

      if(this._useServiceAccount && !this._token){
        await this.getAccessToken();
      }
      this._modelId = modelId;
      await this.getAccountDimensions();
    }
    setUseServiceAccount(useServiceAccount){
      this._useServiceAccount=useServiceAccount;
    }

    async setOAuthParameters(tokenURL,clientID,secret){
      this._tokenURL=tokenURL;
      this._clientID = clientID;
      this._secret = secret;
      this._useServiceAccount = true;
      let token = await this.getAccessToken();
      await this.getCSRFToken();
      return token;

    }

    setChunkSize(chunkSize) {
      this._chunkSize = chunkSize;
    }
    getChunkSize() {
      return this._chunkSize;
    }
    getCurrentBatchStartingNumber() {
      return this._currentBatch;
    }

    //retrieve the data in the CSV file
    getData(sheetName) {
      return this._data[sheetName];
    }

    getUploadResult() {
      return this._uploadResult;
    }
    getTotalRows(sheetName) {
      let totalRows = 0;
      if (sheetName.length > 0) {
        totalRows = this._data[sheetName].length;
      } else {
        totalRows = this._data[this._sheetNames[0]].length;
      }
      return totalRows;
    }
    getSheetNames() {
      return this._sheetNames;
    }

    async connectedCallback() {
      if(this._useServiceAccount){
        if(this._tokenURL.length>0 && this._baseURL.length>0 && this._clientID.length>0 && this._secret.length>0){
          if(this._token.length ===0){
            await this.getAccessToken();
          }
            await this.getCSRFToken();
        }
      }else{
        await this.getCSRFToken();
      }

      await getScriptPromisify(
        "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"
      );
      await getScriptPromisify(
        "https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js"
      );
      this._input.addEventListener("change", (e) => {
        if (e.target.value) {
          this.handleFileSelect(e);
        } else {
          this.dispatch("onCancel"); // Dispatch on Cancel event
        }
      });
      this._remove.addEventListener("click", () => this.handleRemove());
    
    }
  }

  window.customElements.define("com-sap-sample-uploadxls", UploadXLS);
})();
