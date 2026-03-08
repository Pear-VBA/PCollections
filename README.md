# PCollections - Advanced VBA Collection Utilities

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)  
[![VBA](https://img.shields.io/badge/language-VBA-orange.svg)](https://en.wikipedia.org/wiki/Visual_Basic_for_Applications)

**PCollections** is a lightweight VBA library that extends the built-in `Collection` functionality, providing additional methods for searching, indexing, converting, and manipulating Collections in Microsoft Excel, Access, and other VBA-supported environments.

## Features

✅ **Extend** Collections dynamically  
✅ **Search** for items and their indices  
✅ **Convert** Collections to arrays and vice versa  
✅ **Join** Collection elements into strings  
✅ **Find min/max values** in numerical Collections

## Installation

Simply download the `pcol_PCollections.bas` module and import it into your VBA project:

1. Open the **VBA Editor** (`ALT + F11`)
2. Go to **File** → **Import File...**
3. Select `pcol_PCollections.bas` and click **Open**

Alternatively, copy and paste the code into a new module.

## Usage

### 1. Extending a Collection

```vba
Dim Col1 As New Collection, Col2 As New Collection
Col1.Add "A": Col1.Add "B"
Col2.Add "C": Col2.Add "D"
pcol_PCollections.pcol_Extend Col1, Col2  ' Col1 now contains A, B, C, D
```

### 2. Finding Items and Their Indices

```vba
Dim Col As New Collection
Col.Add "Apple"
Col.Add "Banana"
Col.Add "Apple"

Dim Result As Collection
Set Result = pcol_PCollections.pcol_FindAll(Col, "Apple") ' Returns a Collection with both "Apple" items

Dim Index As Long
Index = pcol_PCollections.pcol_FindIndex(Col, "Banana") ' Returns 2

Dim LastIndex As Long
LastIndex = pcol_PCollections.pcol_FindLastIndex(Col, "Apple") ' Returns 3
```

### 3. Converting Between Arrays and Collections

```vba
Dim Arr As Variant
Arr = Array("X", "Y", "Z")

Dim Col As Collection
Set Col = pcol_PCollections.pcol_FromArray(arr)  ' Converts array to Collection

Dim BackToArray As Variant
BackToArray = pcol_PCollections.pcol_ToArray(Col) ' Converts Collection back to array
```

### 4. Checking for Existence of Items or Keys

```vba
Dim Exists As Boolean
Exists = pcol_PCollections.pcol_ItemExists(Col, "X") ' Returns True if "X" is in the Collection

Dim IsKeyExists As Boolean
IsKeyExists = pcol_PCollections.pcol_KeyExists(Col, 1) ' Returns True if a key exists
```

### 5. Working with Collections

```vba
Dim Col As New Collection
Col.Add 10
Col.Add 20
Col.Add 5

Debug.Print pcol_PCollections.pcol_Max(Col)  ' Prints 20
Debug.Print pcol_PCollections.pcol_Min(Col)  ' Prints 5
Debug.Print pcol_PCollections.pcol_Join(Col, ", ") ' Prints "10, 20, 5"
```

## Available Functions

### Collection Manipulation

`pcol_Extend(Destination As Collection, Source As Collection)`: Merges two Collections.

### Searching & Indexing

`pcol_FindAll(SourceCollection As Collection, Item As Variant)`: Returns a Collection of all matching items.
`pcol_FindIndex(SourceCollection As Collection, Item As Variant)`: Returns the first index of an item.
`pcol_FindLastIndex(SourceCollection As Collection, Item As Variant)`: Returns the last index of an item.
`pcol_ItemExists(SourceCollection As Collection, Item As Variant)`: Checks if an item exists.
`pcol_KeyExists(SourceCollection As Collection, Key As Variant)`: Checks if a key exists in a Collection.

### Conversion

`pcol_FromArray(SourceArray As Variant)`: Converts an array to a Collection.
`pcol_ToArray(SourceCollection As Collection)`: Converts a Collection to an array.

### String Operations

`pcol_Join(SourceCollection As Collection, Delimiter As String)`: Joins Collection items into a string.
`pcol_Split(Expression As String, Delimiter As String)`: Splits a string into a Collection.

### Mathematical Operations

`pcol_Max(SourceCollection As Collection)`: Returns the maximum value.
`pcol_Min(SourceCollection As Collection)`: Returns the minimum value.

## Why Use PCollections?

✅ Enhance VBA Collection handling with additional utilities

✅ Reduce repetitive code with built-in search and conversion functions

✅ Improve performance for working with large datasets in Excel, Access, or Outlook

✅ Easy to integrate – just import and start using

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Feel free to submit issues or pull requests to improve this library.

## Keywords (for search optimization)

VBA, Collection, Array, Utilities, Find, Min, Max, Join, Excel VBA, Microsoft Office, Outlook VBA, Access VBA, Search, Split, Extend, Custom Collection
