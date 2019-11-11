yii2-word
========

Helper and wrapper class for PHPWord for simple communication.

## Installation

Either run

```
$ php composer.phar require natmars/yii2-word "@dev"
```

or add

```
"natmars/yii2-word": "@dev"
```

## Creating a multi-line file using the word template ##

### 1. Creating Your First Template

Create a Microsoft Word file. 

Use `${variableName}` to set the variables. Use `${tableName}`, `${itemName}` to generate the table (the value of `${tableName}` is set to empty, but the values ​​of `${itemName}` are set to match the data in the array). Looks like this:

![Word Template Screenshot](/example/templates/template.png?raw=true)

### 2. Saving template

Set the file name as `$fileName`, put the file in the template directory `$reportTemplatesFileDir` and determine the location to save the finished file `$reportLog`

### 3. Generating the report 

```php
use natmars\word\Word;

$templateName = $reportTemplatesFileDir . $fileName;
$newFileName = $reportLog . $fileName;
        
$phpWord = new Word();

// create a file using a template
$phpWord->saveFromMultiLineTemplate($templateName, $newFileName, [
    'variableName' => 'variableValue',
    'tableName' => [
        ['itemName' => 'itemValue1'],
        ['itemName' => 'itemValue2'],
        ['itemName' => 'itemValue3'],
    ],
]);

// download file
\Yii::$app->response->sendFile($newFileName, $fileName);
```

### 4. Check your output report document

Your output should look something like below:

![Output Document Screenshot](/example/templates/output.png?raw=true)

##
Run for example:
```
\natmars\example\Test::generate();
```