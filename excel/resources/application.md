
# Application

Represents the Excel application that manages the workbook. Updated today in the room.

## [Properties](#getter-examples)

| Property | Type | Description |
| --- | --- | --- |
| calculationMode | string | Returns the calculation mode used in the workbook. Read-only. Possible values are: `Automatic` Excel controls recalculation. |
| NewField | int | Returns nothing |

## Relationships

None

## Methods

| Method | Return Type | Description |
| --- | --- | --- |
| [calculate(calculationType: string)](#calculatecalculationtype-string) | void | Recalculate all currently opened workbooks in Excel. |
| [load(param: object)](#loadparam-object) | void | Fills the proxy object created in JavaScript layer with property and object values specified in the parameter. |

## API Specification

### calculate(calculationType: string)

Recalculate all currently opened workbooks in Excel.

#### Syntax

```javascript
applicationObject.calculate(calculationType);
```

#### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| calculationType | string | Specifies the calculation type to use. Possible values are: `Recalculate` Default-option. Performs normal calculation by calculating all the formulas in the workbook.,`Full` Forces a full calculation of the data.,`FullRebuild` Forces a full calculation of the data and rebuilds the dependencies. |

#### Returns

void

#### Examples

```javascript
var ctx = new Excel.RequestContext();

ctx.workbook.application.calculate('Full');

ctx.executeAsync();
```

[Back](#methods)

### load(param: object)

Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax

```javascript
object.load(param);
```

#### Parameters

| Parameter | Type | Description |
| --- | --- | --- |
| param | object | Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object. |

#### Returns

void

#### Examples

[Back](#methods)

### Getter Examples

```javascript
var ctx = new Excel.RequestContext();

var application = ctx.workbook.application;

application.load(calculationMode);

ctx.executeAsync().then(function() {
    Console.log(application.calculationMode);

});
```

[Back](#properties)
