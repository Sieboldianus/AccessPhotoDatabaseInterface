Access Photo Database Interface
=============
Tool to read & import lists of Photo Metadata (*.txt) to MS Access and perform simple Time-Aggregation, Conversion and Export tasks.

**A Photo Database Interface** that helps importing lists of photo metadata to Access for simple analysis and conversion tasks. Tested for up to 5M photo entries.

![photo Database Interface](interface.png?raw=true)

## Motivation

Automation of reoccuring photo metadata processing tasks for testing & visualization. This tool is also used to calculate Tag Statistics for the generation of [tag maps](https://www.flickr.com/photos/64974314@N08/albums/72157628868173205).

## Code Example

The following code counts and aggregates unique users (or photos) per day and exports list into a new table. This code can be used as a template for preparing other reoccuring MySQL queries.

```vb
    Sub UniqueDay_statistics()
		Dim db As DAO.Database
		Dim qdf As DAO.QueryDef
		Dim sSQL As String
		Set form1 = Forms("Database Tools")
		Set db = CurrentDb

		tablename = form1.Text78.Value
		feldname = "TimeStamp"
		feldname_Countdistinct = form1.Text90.Value
		output_name = tablename & "_perUNIQUEDAY"

		If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_name & "'")) Then
			MsgBox "Query " & output_name & ": an object already exists with this name, using this one instead."
		Else
			CurrentDb.CreateQueryDef output_name, "SELECT * FROM taglist_templ"
		End If

		DoCmd.SetWarnings False
		Set qdf = db.QueryDefs(output_name)
	 
		sSQL = " SELECT UNIQUEDAY, COUNT([" & feldname_Countdistinct & "]) AS " & feldname_Countdistinct & "_COUNT"
		sSQL = sSQL & " FROM ( SELECT DISTINCT Format(" & tablename & ".[" & feldname & "],'mm/dd/yyyy') AS UNIQUEDAY, [" & feldname_Countdistinct & "] FROM " & tablename & ") AS TBL_tmp"
		sSQL = sSQL & " GROUP BY UNIQUEDAY"
		qdf.SQL = sSQL

		DoCmd.OpenQuery output_name
		DoCmd.SetWarnings True
			  
		Set qdf = Nothing
		Set db = Nothing
		DoCmd.OpenQuery output_name, acViewNormal, acReadOnly
	End Sub
```

## Installation

* Download [importdata_templ_V4-1.accdb](importdata_templ_V4-1.accdb)
* Can be used with free [2016 Access Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=50040)

## Contributors

* todo: future goals

## Built With
This project includes and makes use of several other projects/libraries/frameworks:

>[*RegExprReplace*](http://www.experts-exchange.com/articles/Programming/Languages/Visual_Basic/Using-Regular-Expressions-in-Visual-Basic-for-Applications-and-Visual-Basic-6.html) - by Patrick G. Matthews
>> Function for using Regular Expressions to parse a string, and replace parts of the string matching the specified pattern with another string.
>> Optionally used to clean photo metadata

## License

GNU GPLv3

## Changelog

* 2017-04-07: [Access Photo Database Interface V4-1](importdata_templ_V4-1.accdb)

	* Initial commit
	* Formatted code and added readme

[//]: # (Readme formatting based on https://gist.github.com/PurpleBooth/109311bb0361f32d87a2) 
