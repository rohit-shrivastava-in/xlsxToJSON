# xlsxToJSON

An angular directive to read XLSX / Spreadsheet / Workbook (Excel Spreadsheet) and convert into `JSON`

# installation

### NPM
```
	$ npm install angular
	$ npm install angular-js-xlsx
	$ npm install xlsx
```
### Bower
```
	$ bower install angular
	$ bower install angular-js-xlsx
	$ bower install js-xlsx
```

# Project integration

# Imports
In the browser, just add a script tag:
```
	<script type="text/javascript" src="node_modules/angular/angular.min.js"></script>
	<script type="text/javascript" src="node_modules/angular-js-xlsx/angular-js-xlsx.js"></script>
	<script type="text/javascript" src="node_modules/angular-js-xlsx/node_modules/xlsx/dist/xlsx.core.min.js"></script>
	<script src="node_modules/xlsx/xlsx.js"></script>
```
  
  
# Module
```javascript
	angular.module('yourApp', ['angular-js-xlsx']);
```

# Directive to read file
```javascript
<js-xls onread="read" onerror="error"></js-xls>
```
read file data
```javascript
$scope.read = function (workbook) {
   /* DO SOMETHING WITH workbook HERE */
   console.log(workbook);
}
```
handle unwated errors
  
```javascript
$scope.error = function (e) {
    /* DO SOMETHING WHEN ERROR IS THROWN */
    console.log(e);
}
```
  
# Working with the Workbook

```javascript
    var first_sheet_name = workbook.SheetNames[0];
    var address_of_cell = 'A1';
```

Get worksheet
```javascript
    var worksheet = workbook.Sheets[first_sheet_name];
```

Find desired cell
```javascript
    var desired_cell = worksheet[address_of_cell];
```

Get the value
```javascript
    var desired_value = (desired_cell ? desired_cell.v : undefined);
```

# Utilities

Utilities are available in the `XLSX` object and are described in the Utility Functions section:

```javascript
    var X = XLSX;
```
  
# Convert data into 

### CSV
`sheet_to_csv` generates delimiter-separated-values output.
```javascript
var to_csv = function to_csv(workbook) {
	var result = [];
	workbook.SheetNames.forEach(function(sheetName) {
		var csv = X.utils.sheet_to_csv(workbook.Sheets[sheetName]);
		if(csv.length){
			result.push("SHEET: " + sheetName);
			result.push("");
			result.push(csv);
		}
	});
	return result.join("\n");
};
  ```

### Formulae 
`sheet_to_formulae` generates a list of the formulae (with value fallbacks).
  
```javascript
    var to_fmla = function to_fmla(workbook) {
		var result = [];
		workbook.SheetNames.forEach(function(sheetName) {
			var formulae = X.utils.get_formulae(workbook.Sheets[sheetName]);
			if(formulae.length){
				result.push("SHEET: " + sheetName);
				result.push("");
				result.push(formulae.join("\n"));
			}
		});
		return result.join("\n");
	};
```

`sheet_to_html` generates HTML output.

```javascript
    var to_html = function to_html(workbook) {
		HTMLOUT.innerHTML = "";
		workbook.SheetNames.forEach(function(sheetName) {
			var htmlstr = X.write(workbook, {sheet:sheetName, type:'binary', bookType:'html'});
			HTMLOUT.innerHTML += htmlstr;
		});
		return "";
	};
```
#### JSON

`sheet_to_json` converts a worksheet object to an array of JSON objects.

```javascript
    var to_json = function (workbook) {
        var result = {};
        workbook.SheetNames.forEach(function(sheetName) {
            var roa = X.utils.sheet_to_json(workbook.Sheets[sheetName]);
            if(roa.length) result[sheetName] = roa;
        });
        return result;
    };
 ```
#### Excel data input
  Shee1

| Name | Email |
| ------ | ------ |
| abc | abc@gmail.com |
| xyz | xyzl@gmail.com |
  
  #### Output
  
```javascript
 {
  "Sheet1": [
    {
      "Name": "abc",
      "email": "abc@gmail.com"
    },
    {
      "Name": "xyz",
      "email": "xyz@houzlook.com"
    }
  ]
}
 ```