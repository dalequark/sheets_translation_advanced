/* This script is designed to run attached to a Google Sheets. 
Create a new sheet, then select Tools -> Script Editor.

You'll also need these "Script Properties":

bucket: YOUR_BUCKET_NAME
targetLang: "es" // some language code
srcLang: "en" // some language code
projectId: YOUR_PROJECT_ID
glossaryId: // This will be set in code

*/

function onTranslationsEdit(e) {
    // This function should be called in an external onEdit trigger.

    // Only do anything if we changed the translations sheet
    if (e.range.getSheet().getName() != "Translations") return;

    // Must be a change in the source language row, too:
    if (e.range.getColumn() != 1) return;

    if (!e.value || e.value == e.oldValue) return;

    var glossaryId = PropertiesService.getScriptProperties().getProperty('glossaryId');
    var srcLang = PropertiesService.getScriptProperties().getProperty('srcLang');
    var targetLang = PropertiesService.getScriptProperties().getProperty('targetLang');

    var translation = translateWithGlossary([e.value], srcLang, targetLang, glossaryId)

    e.range.getSheet().getRange(e.range.getRow(), e.range.getColumn() + 1).setValue(translation);

}

function updateGlossary() {
    // Takes the glossary in the "Glossary" sheet and translates
    // all the words in the "Translations" sheet with this new glossary.
    alert("Updating glossary and translations. This could take a moment.");

    const [srcLang, targetLang, glossary] = getGlossaryFromSheet();

    PropertiesService.getScriptProperties().setProperty('srcLang', srcLang);
    PropertiesService.getScriptProperties().setProperty('targetLang', targetLang);

    Logger.log("Converting from " + srcLang + " to " + targetLang);
    Logger.log("Glossary is :");
    Logger.log(glossary);
    // Uploads this glossary to gcs bucket
    var glossaryName = csvFromDict(glossary, "glossary");
    if (!glossaryName) {
        alert("Could not upload glossary to GCS. Sorry :(");
        return;
    }
    const glossaryId = 'g' + glossaryName.split(".")[0];

    PropertiesService.getScriptProperties().setProperty('glossaryId', glossaryId);

    const operationName = createGlossary(glossaryName, glossaryId, srcLang, targetLang);

    while (!checkGlossaryCreationStatus(operationName)[0]) {
        Utilities.sleep(1000);
    }

    const status = checkGlossaryCreationStatus(operationName)[1];
    if (status != "SUCCEEDED") {
        const errorMsg = "Error creating glossary: " + status;
        alert(errorMsg);
        Logger.log(errorMsg);
        return;
    }
    Logger.log("Successfully created glossary");
    updateTranslations(srcLang, targetLang, glossaryId);
}

function updateTranslations(srcLang, targetLang, glossaryId) {
    // Given a glossaryId, pulls the words to be 
    // translated from the sheet and writes their
    // translations.

    var srcText = getSrcText();
    var translations = translateWithGlossary(srcText, srcLang, targetLang, glossaryId);
    writeTranslations(translations);
}

function testUpdateTranslations() {
    updateTranslations("en", "fr", "g15026");
}

function getSrcText() {
    // Gets sentences to be translated from the "Translations" sheet.
    // Returns an array sentences in the source language to be
    // translated.

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Translations");
    const allData = sheet.getDataRange().getValues();
    var srcText = allData.map(function (row) {
        return row[0];
    });
    // Get rid of the header row.
    srcText.shift();
    return srcText;
}

function testGetSrcText() {
    Logger.log(getSrcText());
}

function writeTranslations(translationsArray) {
    // Writes the words in the translationsArray to the second
    // column in the "Translations" sheet.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Translations");
    Logger.log("Array length: " + translationsArray.length);
    const rows = sheet.getRange(2, 2, translationsArray.length);
    for (var i = 0; i < translationsArray.length; i++) {
        var cell = sheet.getRange(i + 2, 2);
        cell.setValue(translationsArray[i]);
    }
}

function getGlossaryFromSheet() {
    // Pulls the glossary from the current spreadsheet.
    // Returns a (srcLang, targetLang, glossary) array.

    const glossarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Glossary");
    const allData = glossarySheet.getDataRange().getValues();

    try {
        const srcLang = allData[0][0];
        const targetLang = allData[0][1];
        if (!srcLang.length || !targetLang.length) {
            throw "Empty src or dest lang";
        }
    } catch (e) {
        alert("Need to set a valid source and destination language in the top two cells " + e);
        return;
    }

    const glossary = {};
    for (rowIdx in allData) {
        var row = allData[rowIdx];
        if (row.length < 2) continue;
        if (row[0] && row[1]) {
            glossary[row[0]] = row[1];
        }
    }
    return [srcLang, targetLang, glossary];
}

function alert(text) {
    SpreadsheetApp.getUi().alert(text);
}

/* -------- Translation V3 Utils ----- */

function translateWithGlossary(textArray, srcLang, targetLang, glossaryId) {
    // Translates text arrat from the given src to target lang using a glossary.

    const projectId = PropertiesService.getScriptProperties().getProperty("projectId");

    var url = "https://translation.googleapis.com/v3beta1/projects/PROJECTID/locations/us-central1:translateText"
        .replace("PROJECTID", projectId);

    var data = {
        source_language_code: srcLang,
        target_language_code: targetLang,
        contents: textArray,
        glossary_config: {
            glossary: "projects/PROJECTID/locations/us-central1/glossaries/GLOSSARYID"
                .replace("PROJECTID", projectId)

                .replace("GLOSSARYID", glossaryId)
        }
    };

    var response = UrlFetchApp.fetch(url, {
        method: "POST",
        contentType: "application.json; charset=utf-8",
        payload: JSON.stringify(data),
        headers: {
            Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        }
    });

    if (response.getResponseCode() != 200) {
        Logger.log("Error translating text: " + response.getResponseCode());
        return false;
    }

    response = JSON.parse(response.getContentText());
    console.log("Got translation response: ", response);
    return response["glossaryTranslations"].map(function (x) {
        return x["translatedText"];
    });


}

function testTranslateWithGlossary() {
    Logger.log(translateWithGlossary(["This is a test for cats",
        "This is a test for dogs"],
        "en", "fr", "g15026"));
}

function createGlossary(gcsFilename, glossaryId, sourceLang, targetLang) {
    // Creates a glossary from an existing csv file
    // in a gcs location and gives it the id glossaryId.
    // gcsFilename should be just the filename, not the path.
    // Returns false on error or the name of the long-running cloud operation.

    Logger.log("Create glossary from csv...");
    const projectId = PropertiesService.getScriptProperties().getProperty("projectId");
    const bucket = PropertiesService.getScriptProperties().getProperty("bucket");
    var url = "https://translation.googleapis.com/v3beta1/projects/PROJECTID/locations/us-central1/glossaries"
        .replace("PROJECTID", projectId);

    var data = {
        name: "projects/PROJECTID/locations/us-central1/glossaries/GLOSSARYID"
            .replace("PROJECTID", projectId)

            .replace("GLOSSARYID", glossaryId),
        language_pair: {
            source_language_code: sourceLang,
            target_language_code: targetLang,
        },
        input_config: {
            gcs_source: {
                input_uri: 'gs://BUCKETNAME/GLOSSARYFILE'
                    .replace("BUCKETNAME", bucket)
                    .replace("GLOSSARYFILE", gcsFilename)
            }
        }
    };

    var response = UrlFetchApp.fetch(url, {
        method: "POST",
        contentType: "application.json; charset=utf-8",
        payload: JSON.stringify(data),
        headers: {
            Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        }
    });

    if (response.getResponseCode() != 200) {
        var errorMsg = "Error creating glossary: " + response.getResponseCode();
        alert(errorMsg);
        Logger.log(errorMsg);
        return false;
    }
    // This returns "RUNNING" until the glossary is done being created. //
    response = JSON.parse(response.getContentText());
    const operationName = response["name"];

    Logger.log("Creating glossary, got response " + response["metadata"]["state"]);
    return operationName;
}


function checkGlossaryCreationStatus(operationName) {
    // Returns glossary status as a tuple of "done" (true/false) 
    // and "status" (RUNNING/SUCCESS/etc)
    const url = "https://translation.googleapis.com/v3beta1/" + operationName;
    Logger.log("operation name is " + operationName);
    var response = UrlFetchApp.fetch(url, {
        method: "GET",
        contentType: "application.json",
        headers: {
            Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        }
    });
    // This returns "RUNNING" until the glossary is done being created. //
    response = JSON.parse(response.getContentText());
    if (response["metadata"]["state"] == "FAILED") {
        Logger.log("Failed to create glossary");
        Logger.log(response);
    }
    return [response["done"], response["metadata"]["state"]];
}


function testCreateGlossary() {
    const operation = createGlossary("18022.csv", "myfirstglossary", "en", "es");
    for (var i = 0; i < 5; i++) {
        Logger.log(checkGlossaryCreationStatus(operation));
    }
}

/* -------- GCS Storage Utils -------- */
function csvFromDict(glossaryDict) {
    // Uploads a dictionary to a GCS bucket as a file
    // called <glossaryName>.csv. Returns 0 for failure,
    // or the (randomly generated) name of the csv created.

    const glossaryName = Math.ceil(Math.random() * Math.exp(10, 10)).toString() + ".csv";

    Logger.log("Trying to create glossary named " + glossaryName);
    // Construct a csv from the passed dict
    var csv = "";
    for (var key in glossaryDict) {
        csv += key + "," + glossaryDict[key] + "\n";
    }

    const bucket = PropertiesService.getScriptProperties().getProperty("bucket");
    var url = 'https://www.googleapis.com/upload/storage/v1/b/BUCKET/o?uploadType=media&name=FILE'
        .replace("BUCKET", bucket)
        .replace("FILE", encodeURIComponent(glossaryName));

    var response = UrlFetchApp.fetch(url, {
        method: "POST",
        contentLength: csv.length,
        contentType: "text/csv",
        payload: csv,
        headers: {
            Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        }
    });
    Logger.log("Got response " + response.getResponseCode());
    if (response.getResponseCode() != 200) {
        Logger.log("Error uploading glossary to GCS: " + response.getResponseCode());
        return false;
    }

    response = JSON.parse(response.getContentText());
    Logger.log("Tried to create gcs csv file " + glossaryName);
    return glossaryName;
}

function listBucket() {
    // Lists bucket files

    const bucket = PropertiesService.getScriptProperties().getProperty("bucket");
    var url = "https://www.googleapis.com/storage/v1/b/BUCKET_NAME/o"
        .replace("BUCKET_NAME", bucket);

    var response = UrlFetchApp.fetch(url, {
        method: "GET",
        headers: {
            Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        }
    });
    response = JSON.parse(response.getContentText());
    var objects;
    if (response["items"]) {
        objects = response["items"].map(function (bucket) { return bucket.name; }).join("\n");
    }
    else {
        objects = "No objects found in bucket";
    }
    return objects;
}

/*------------ Menu Tools ---------------*/
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Translation')
        .addItem('Update Glossary', 'updateGlossary')
        .addToUi();
}
