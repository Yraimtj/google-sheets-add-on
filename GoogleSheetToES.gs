/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
const onOpen = function (e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Send To Cluster...", "showPushDataSidebar")
    .addToUi();
};

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
const onInstall = function (e) {
  onOpen(e);
};

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
const showPushDataSidebar = function () {
  let ui = HtmlService.createTemplateFromFile("ConnectionDetailsSidebar")
    .evaluate()
    .setTitle("Send Data To Cluster");
  SpreadsheetApp.getUi().showSidebar(ui);
};

/**
 * Checks to see if the cluster is accessible by calling /_status
 * Throws an error if the cluster does not return a 200
 *
 * @param {Object} host The set of parameters needed to connect to a cluster.
 */
const checkClusterConnection = function (host) {
  isValidHost(host);
  let url = [
    host.use_ssl ? "https://" : "http://",
    host.host,
    ":",
    host.port,
    "/",
  ].join("");
  let options = getDefaultOptions(host.username, host.password);
  options["muteHttpExceptions"] = true;
  try {
    let resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() != 200) {
      let jsonData = JSON.parse(resp.getContentText());
      if (jsonData.message == "forbidden") {
        throw "The username and/or password is incorrect.";
      }
      throw jsonData.message;
    }
  } catch (e) {
    throw "There was a problem connecting to your cluster. Please the connection details and try again.";
  }
};

const clearData = function () {
  let userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
};

const saveHostData = function (host) {
  isValidHost(host);
  let userProperties = PropertiesService.getUserProperties();
  userProperties.setProperties(host);
};

const getHostData = function () {
  let userProperties = PropertiesService.getUserProperties();
  let data = userProperties.getProperties();
  return {
    host: data["host"],
    port: data["port"],
    use_ssl:
      typeof data["use_ssl"] == "string"
        ? data["use_ssl"] == "true"
        : data["use_ssl"],
    username: data["username"],
    password: data["password"],
    was_checked: data["was_checked"],
  };
};

/**
 * Returns a clean name to use as an index based on the sheet name
 *
 */
const getSheetName = function () {
  try {
    let name = SpreadsheetApp.getActiveSheet().getName();
    return name.replace(/[^0-9a-zA-Z]/g, "_").toLowerCase();
  } catch (e) {
    return "";
  }
};

/**
 * Highlights the cells in the A1 range
 * @param {String} a1_range A1 notation for the cells to highlight - required.
 */
const highlightData = function (a1_range) {
  let sheet = SpreadsheetApp.getActiveSheet();
  try {
    let range = sheet.getRange(a1_range);
    if (a1_range.length == 3) {
      sheet.setActiveSelection(
        sheet.getRange(range.getRow(), range.getColumn(), range.getHeight())
      );
    } else {
      sheet.setActiveSelection(range);
    }
  } catch (e) {
    throw "The range entered was invalid. Please verify the range entered.";
  }
};

/**
 * Gets the default locations for headers and data, namely the first row
 * and all other rows.
 */
const getDefaultRange = function () {
  try {
    let sheet = SpreadsheetApp.getActiveSheet();
    let data_range = sheet.getRange(
      1,
      1,
      sheet.getLastRow(),
      sheet.getLastColumn()
    );
    return data_range.getA1Notation();
  } catch (e) {
    throw "There is no data in the sheet.";
  }
};

const getSelectedRange = function () {
  try {
    let sheet = SpreadsheetApp.getActiveSheet();
    return sheet.getActiveRange().getA1Notation();
  } catch (e) {
    throw "No range selected.";
  }
};

/**
 * Attempts to validate that the data in each column is the same format.
 * If something isn't the same, it adds a note to the sheet and throws an
 * exception.
 */
const validateData = function (new_value) {
  let sheet = SpreadsheetApp.getActiveSheet();
  let range = null;
  try {
    range = sheet.getRange(new_value);
  } catch (e) {
    throw "There is no data in the sheet.";
  }
  clearNotes();
  let start_row = parseInt(range.getRow()) + 1;
  let start_col = parseInt(range.getColumn());
  let formats = range.getNumberFormats();
  formats.shift();
  let header_formats = formats.shift();
  for (let r in formats) {
    for (let c in formats[r]) {
      if (formats[r][c] != header_formats[c]) {
        let note_row = start_row + 1 + parseInt(r);
        let note_col = start_col + parseInt(c);
        let cell = sheet.getRange(note_row, note_col);
        cell.setNote(
          "Not the same format as first row. This may cause data to not be inserted into your cluster. ~SpreadsheetToES"
        );
        throw "Not all data formats are the same. See the note in the sheet.";
      }
    }
  }
};

/**
 * Attempts to clear only the notes that we've made
 */
const clearNotes = function () {
  let sheet = SpreadsheetApp.getActiveSheet();
  let notes_range = sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .getNotes();
  for (let r in notes_range) {
    for (let c in notes_range[r]) {
      if (
        notes_range[r][c] &&
        notes_range[r][c].indexOf("~SpreadsheetToES") >= 0
      ) {
        sheet.getRange(1 + parseInt(r), 1 + parseInt(c)).clearNote();
      }
    }
  }
};

/**
 * Pushes data from the spreadsheet to the cluster.
 *
 * @param {Object} host The set of parameters needed to connect to a cluster - required.
 * @param {String} index The index name - required.
 * @param {String} index_type The index type - required.
 * @param {String} template The name of the index template to use.
 * @param {String} header_range The A1 notion of the header row.
 * @param {String} data_range The A1 notion of the data rows.
 */

const pushDataToCluster = function (
  index = "test-default",
  index_type = "default-type",
  template,
  data_range_a1
) {
  let host = getHostData();
  isValidHost(host);

  let doc_id_range_a1 = "A:A";

  checkInput(index, index_type, template, data_range_a1);

  let [first_column, first_row, last_column, last_row] =
    data_range_a1.match(/[a-zA-Z]+|[0-9]+/g);

  let [data, data_range] = getDataRange(data_range_a1);

  let doc_id_data = getDocIdData(doc_id_range_a1, data_range);

  let headers;
  if (first_row === "1") {
    headers = data.shift();
  } else {
    headers_range = `${first_column}1:${last_column}1`;
    try {
      headers = SpreadsheetApp.getActiveSheet()
        .getRange(headers_range)
        .getValues()[0];
    } catch (e) {
      throw "The document data range entered was invalid. Please verify the range entered.";
    }
  }

  for (let i in headers) {
    if (!headers[i]) {
      throw "Document key name cannot be empty. Please make sure each cell in the document key names range has a value.";
    }
    headers[i] = headers[i].replace(/[^0-9a-zA-Z]/g, "_"); // clean up the column names for index keys

    headers[i] = headers[i].toLowerCase();
    if (!headers[i]) {
      throw "Document key name cannot be empty. Please make sure each cell in the document key names range has a value.";
    }
  }

  let bulkList = [];
  if (template) {
    createTemplate(host, index, template);
  }
  let did_send_some_data = false;
  for (let r = 0; r < data.length; r++) {
    let row = data[r];
    let toInsert = {};
    for (let c = 0; c < row.length; c++) {
      if (row[c]) {
        toInsert[headers[c]] = row[c];
      }
    }
    if (Object.keys(toInsert).length > 0) {
      if (doc_id_data) {
        if (!doc_id_data[r][0]) {
          throw "Missing document id for data row: " + (r + 1);
        }
        bulkList.push(
          JSON.stringify({
            update: {
              _index: index,
              _type: "_doc",
              _id: doc_id_data[r][0],
              retry_on_conflict: 3,
            },
          })
        );
        bulkList.push(
          JSON.stringify({
            doc: toInsert,
            detect_noop: true,
            doc_as_upsert: true,
          })
        );
      } else {
        bulkList.push(
          JSON.stringify({ index: { _index: index, _type: "_doc" } })
        );
        bulkList.push(JSON.stringify(toInsert));
      }
      did_send_some_data = true;
      // Don't hit the UrlFetchApp limits of 10MB for POST calls.
      if (bulkList.length >= 2000) {
        postDataToES(host, bulkList.join("\n") + "\n");
        bulkList = [];
      }
    }
  }

  if (bulkList.length > 0) {
    postDataToES(host, bulkList.join("\n") + "\n");
    did_send_some_data = true;
  }
  if (!did_send_some_data) {
    throw "No data was sent to the cluster. Make sure your document key name and value ranges are valid.";
  }
  return [
    host.use_ssl ? "https://" : "http://",
    host.host,
    ":",
    host.port,
    "/",
    index,
    "/_search",
  ].join("");
};

const checkInput = function (index, index_type, template, data_range_a1) {
  if (index.indexOf(" ") >= 0) {
    throw "Index should not have spaces.";
  }

  if (index_type.indexOf(" ") >= 0) {
    throw "Index type should not have spaces.";
  }

  if (template && template.indexOf(" ") >= 0) {
    throw "Template name should not have spaces.";
  }

  if (!data_range_a1) {
    throw "Document data range cannot be empty.";
  }
};

const getDataRange = function (data_range_a1) {
  let data_range = null;
  try {
    data_range = SpreadsheetApp.getActiveSheet().getRange(data_range_a1);
  } catch (e) {
    throw "The document data range entered was invalid. Please verify the range entered.";
  }
  let data = data_range.getValues();

  if (data.length <= 0) {
    throw "No data to push.";
  }
  return [data, data_range];
};

const getDocIdData = function (doc_id_range_a1, data_range) {
  let doc_id_data;
  let doc_id_range = null;
  try {
    doc_id_range = SpreadsheetApp.getActiveSheet().getRange(doc_id_range_a1);
  } catch (e) {
    throw "The document id column entered was invalid. Please verify the id column entered.";
  }

  if (first_row === "1") {
    doc_id_range = doc_id_range.offset(
      data_range.getRow(),
      0,
      data_range.getHeight() - 1
    );
  } else {
    doc_id_range = doc_id_range.offset(
      data_range.getRow() - 1,
      0,
      data_range.getHeight()
    );
  }
  doc_id_data = doc_id_range.getValues();
  return doc_id_data;
};
/**
 * Delete data in the cluster and also clear it in the spreadsheet.
 *
 * @param {Object} host The set of parameters needed to connect to a cluster - required.
 * @param {String} index The index name - required.
 * @param {String} index_type The index type - required.
 * @param {String} template The name of the index template to use.
 * @param {String} header_range The A1 notion of the header row.
 * @param {String} data_range The A1 notion of the data rows.
 */
const deleteRow = function (
  index = "test-default",
  index_type = "default-type",
  template,
  data_range_a1
) {
  let host = getHostData();
  isValidHost(host);

  let doc_id_range_a1 = "A:A";

  checkInput(index, index_type, template, data_range_a1);

  let [first_column, first_row, last_column, last_row] =
    data_range_a1.match(/[a-zA-Z]+|[0-9]+/g);

  let [data, data_range] = getDataRange(data_range_a1);

  let doc_id_data = getDocIdData(doc_id_range_a1, data_range);

  if (first_row === "1") {
    throw "Can't Delete first row(headers) of this sheet.";
  }

  let bulkList = [];

  let did_send_some_data = false;
  for (let r = 0; r < data.length; r++) {
    if (doc_id_data) {
      if (!doc_id_data[r][0]) {
        throw "Missing document id for data row: " + (r + 1);
      }
      bulkList.push(
        JSON.stringify({
          delete: {
            _index: index,
            _type: "_doc",
            _id: doc_id_data[r][0],
            retry_on_conflict: 3,
          },
        })
      );
    } else {
      throw "Document data id range cannot be empty...";
    }
    did_send_some_data = true;
    // Don't hit the UrlFetchApp limits of 10MB for POST calls.
    if (bulkList.length >= 2000) {
      postDataToES(host, bulkList.join("\n") + "\n");
      bulkList = [];
    }
  }

  if (bulkList.length > 0) {
    postDataToES(host, bulkList.join("\n") + "\n");
    did_send_some_data = true;
  }
  if (!did_send_some_data) {
    throw "No data was sent to the cluster. Make sure your document key name and value ranges are valid.";
  }
  clearDataInRange(data_range_a1);
  return data_range_a1;
};

/**
 * Remove column in the spreadsheet and field in the cluster.
 */
const deleteColumn = function (index, col_range) {
  let host = getHostData();
  isValidHost(host);

  if (index.indexOf(" ") >= 0) {
    throw "Index should not have spaces.";
  }

  if (!col_range) {
    throw "Document data range cannot be empty.";
  }

  let [first_column, first_row, last_column, last_row] =
    data_range_a1.match(/[a-zA-Z]+|[0-9]+/g);

  if (first_column === "A") {
    throw "Can't delete id Column. Please verify the range entered. ";
  }

  let firstRowIsALetter = /[a-zA-Z]/.test(first_row);

  let firstRowNotSelected = (first_row !== "1") & !firstRowIsALetter;

  if (firstRowNotSelected) {
    throw "The document data range entered was invalid. Need to select first row of the column or all the column";
  }

  if (firstRowIsALetter) {
    col_range = `${first_column}:${first_row}`;
  } else {
    col_range = `${first_column}:${last_column}`;
  }

  let data_range_col = null;
  try {
    data_range_col = SpreadsheetApp.getActiveSheet()
      .getRange(col_range)
      .getValues();
  } catch (e) {
    throw "The document data range entered was invalid. Please verify the range entered.";
  }

  let headers_name = data_range_col[0];

  updateByQueryRequest(index, headers_name, host);

  clearDataInRange(col_range);

  return col_range;
};

/**
 * Creates a index template if required. If template already exists, it
 * does not update. If not, it uses default_template and the template name
 * to create a new one.
 *
 * @param {Object} host The set of parameters needed to connect to a cluster - required.
 * @param {String} index The index name - required.
 * @param {String} template_name The name of the index template to use - required.
 */
const createTemplate = function (host, index, template_name) {
  let url = [
    host.use_ssl ? "https://" : "http://",
    host.host,
    ":",
    host.port,
    "/_template/",
    template_name,
  ].join("");

  let options = getDefaultOptions(host.username, host.password);
  options["muteHttpExceptions"] = true;
  let resp = null;
  try {
    let resp = UrlFetchApp.fetch(url, options);
  } catch (e) {
    throw "There was an issue creating the template. Please check the names of the template or index and try again.";
  }
  if (resp.getResponseCode() == 404) {
    options = getDefaultOptions(host.username, host.password);
    options.method = "POST";
    default_template.template = index;
    options["payload"] = JSON.stringify(default_template);
    options.headers["Content-Type"] = "application/json";
    options["muteHttpExceptions"] = true;
    resp = null;
    try {
      resp = UrlFetchApp.fetch(url, options);
    } catch (e) {
      throw "There was an issue creating the template. Please check the names of the template or index and try again.";
    }
    if (resp.getResponseCode() != 200) {
      let jsonData = JSON.parse(resp.getContentText());
      throw jsonData.message;
    }
  } else if (resp.getResponseCode() == 200) {
    let jsonResp = JSON.parse(resp.getContentText());
    if (jsonResp[template_name].template) {
      let re = new RegExp(jsonResp[template_name].template);
      if (!re.test(index)) {
        throw (
          "The template specified will only be applied to indices matching the following naming pattern: '" +
          jsonResp[template_name].template +
          "' Please update the template or choose a new name."
        );
      }
    }
  }
};

/**
 * Posts data to the ES cluster using the /_bulk endpoint
 *
 * @param {Object} host The set of parameters needed to connect to a cluster - required.
 * @param {Array} data The data to push in an array of JSON strings - required.
 */
const postDataToES = function (host, data) {
  let url = [
    host.use_ssl ? "https://" : "http://",
    host.host,
    ":",
    host.port,
    "/_bulk",
  ].join("");
  let options = getDefaultOptions(host.username, host.password);
  options.method = "POST";
  options["payload"] = data;
  options.headers["Content-Type"] = "application/x-ndjson";
  options["muteHttpExceptions"] = true;
  let resp = null;
  try {
    resp = UrlFetchApp.fetch(url, options);
  } catch (e) {
    throw "There was an error sending data to the cluster. Please check your connection details and data.";
  }
  if (resp.getResponseCode() != 200) {
    let jsonData = JSON.parse(resp.getContentText());
    if (jsonData.error) {
      if (jsonData.error.indexOf("AuthenticationException") >= 0) {
        throw "The username and/or password is incorrect.";
      }
      throw jsonData.error;
    }
    throw "Your cluster returned an unknown error. Please check your connection details and data.";
  }
};

const updateByQueryRequest = function (index, headers_name, host) {
  let url = [
    host.use_ssl ? "https://" : "http://",
    host.host,
    ":",
    host.port,
    "/",
    index,
    "/_update_by_query?conflicts=proceed",
  ].join("");

  let scripts = "";
  let querys = { bool: { should: [] } };
  headers_name.forEach((headerName) => {
    headerName = headerName.toLowerCase();
    scripts += `ctx._source.remove("${headerName}");`;
    querys["bool"]["should"].push({ exists: { field: headerName } });
  });

  let data = JSON.stringify({ script: scripts, query: querys });

  let options = getDefaultOptions(host.username, host.password);
  options.method = "POST";
  options["payload"] = data;
  options.headers["Content-Type"] = "application/x-ndjson";
  options["muteHttpExceptions"] = true;

  UrlFetchApp.fetch(url, options);
};

const deleteRequest = function (index) {
  let host = getHostData();
  isValidHost(host);

  let url = [
    host.use_ssl ? "https://" : "http://",
    host.host,
    ":",
    host.port,
    "/",
    index,
  ].join("");
  let options = getDefaultOptions(host.username, host.password);
  options.method = "DELETE";
  options.headers["Content-Type"] = "application/x-ndjson";
  options["muteHttpExceptions"] = true;

  UrlFetchApp.fetch(url, options);

  return index;
};

const reindexRequest = function (index, tmp_index) {
  let host = getHostData();
  isValidHost(host);

  let url = [
    host.use_ssl ? "https://" : "http://",
    host.host,
    ":",
    host.port,
    "/_reindex",
  ].join("");
  let data = JSON.stringify({
    source: { index: index },
    dest: { index: tmp_index },
  });

  let options = getDefaultOptions(host.username, host.password);
  options.method = "POST";
  options["payload"] = data;
  options.headers["Content-Type"] = "application/x-ndjson";
  options["muteHttpExceptions"] = true;

  UrlFetchApp.fetch(url, options);

  return [index, tmp_index];
};

/**
 * Helper function to get the default UrlFetchApp parameters
 *
 * @param {String} username The username for basic auth.
 * @param {String} password The password for basic auth.
 */
const getDefaultOptions = function (username, password) {
  let options = {
    method: "GET",
    headers: {},
  };
  if (username) {
    options.headers["Authorization"] =
      "Basic " + Utilities.base64Encode(username + ":" + password);
  }
  return options;
};

/**
 * Helper function to validate the host object
 *
 * @param {Object} host The set of parameters needed to connect to a cluster - required.
 */
const isValidHost = function (host) {
  if (!host) {
    throw "Cluster details cannot be empty.";
  }
  if (!host.host || !host.port) {
    throw "Please enter your cluster host and port.";
  }
  if (host.host == "localhost" || host.host == "0.0.0.0") {
    throw "Your cluster must be externally accessible to use this tool.";
  }
};

/**
 * Helper function to clear the data range
 * in the activeGoogle Sheets
 *
 */
const clearDataInRange = function (data_range_a1) {
  try {
    data_range = SpreadsheetApp.getActiveSheet().getRange(data_range_a1);
  } catch (e) {
    throw "The document data range entered was invalid. Please verify the range entered.";
  }
  data_range.clearContent();
};

/**
 * This is the default template to use. The template ke will
 * be relaced with the index name if required.
 *
 */
let default_template = {
  order: 0,
  template: "", // will be replaced with index name
  settings: {
    "index.refresh_interval": "5s",
    "index.analysis.analyzer.default.type": "standard",
    "index.number_of_replicas": "1",
    "index.number_of_shards": "1",
    "index.analysis.analyzer.default.stopwords": "_none_",
  },
  mappings: {
    dynamic_templates: [
      {
        string_fields: {
          mapping: {
            fields: {
              raw: {
                type: "keyword",
                ignore_above: 256,
              },
            },
            type: "text",
          },
          match_mapping_type: "string",
          match: "*",
        },
      },
    ],
  },
  aliases: {},
};
