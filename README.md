# Google Translate API Advanced from a Google Sheet

If you've ever tried to translate text in a Google Sheet before, 
you probably know there's a built-in function called `GoogleTranslate`
that does just that.

But in this example, we'll use Google's [Translate API Advanced](https://medium.com/google-cloud/improving-machine-translation-with-the-google-translation-v3-api-9dc2676e7fb9)
to add a translations with a glossary to Google Sheets.

This will add a dropdown in the Google Sheets toolbar called "Translation" that will let
you use a glossary you've defined in one sheet to translate another.

## Getting Started

First, create a new Google Sheet. In the top bar, click Tools -> Script Editor.
Paste the code in GlossaryTranslate.gs into the text editor.

Next, you'll need to define some Project Properties. In the code editor, click
File -> Project properties. Under "Script properties," set:

```
bucket: YOUR_BUCKET_NAME
targetLang: "es" // some language code
srcLang: "en" // some language code
projectId: YOUR_PROJECT_ID
glossaryId: // This will be set in code
```

## Linking a GCP Project

Because we're using the Translation API Advanced for this project, you'll need
to associate your Sheet with a GCP project. Create a new GCP project and link
it to your sheet by (within the code editor), clicking Resources -> Cloud Platform project.
Then enter your project number (not id!).

## Creating a Storage Bucket

The Translations API Advanced creates glossaries by using csv's stored in 
storage buckets. For this project, you'll need to [create a new bucket](https://cloud.google.com/storage/docs/creating-buckets)
and set the permissions so that your sheet can write files to it. 
Add the name of this bucket to the Script Properties variable `bucket` 
(no need to include `gs://` in front of the bucket name).


## Adding a trigger

Finally, you'll need to add an onEdit trigger that causes the translate function to
run whenever you make a change to your spreadsheet. In GlossaryTranslate.gs, the 
function we want to run on edit is called `onTranslationsEdit`.

Within the code editor top bar, click Edit -> Current project's triggers. Here,
add the function `onTranslationsEdit` to the triggeer `onEdit`.

