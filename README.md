# Karsten Held's WhizzyVBA
A collection of VBA, VB6 and VBScript classes and tools (MIT license)

- [DictCollection.cls](#dictcollectioncls) - A mix between Scripting.Dictionary and Collection that can emulate both

- [DictCollection.vbs](#dictcollectionvbs) - A stripped-down VBScript version of DictCollection.cls

Read the [Comparison of Collections in VB6 and VBA](#comparison-of-collections-in-vb6-and-vba) to understand how DictCollection can be compared to Scripting.Dictionary and Collection.

## DictCollection.cls

DictCollection.cls is a mix between Scripting.Dictionary and Collection that can emulate both. It stores items with or without keys as an ordered list and has extended functionality for dealing with subcollections. Keys are stored internally as a sorted list.

### Features

- **Optimized for Debugging**: See keys and items in VB6/VBA Watch Window.
- **Compatible**: Can emulate and replace Scripting.Dictionary or VB6/VBA Collection
- **Versatile**: Can be used as Array (ArrayList), Key-Value-Store (Map) and Object Tree (JSON storage)
- **Fast**: Retrieve items by index with array-like speed (faster than Dictionary and Collection)
- **Nonthrowing Design**: Does not throw errors by default (configurable using `dc.ThrowErrors`)
- **Endless Subcollection Chaining**: `dc("key1")("key2")("nonexisting_key")("key3")` always returns DictCollections that can be NONEXISTING, EMPTY or FILLED. Checking if a nested item exists is a one-liner.
- **High Test Coverage**: `.SelfTest()` covers all important functions and emulation features

### How to use

A step-by-step introduction into the basic functionality of DictCollection. See also [Other Collection Functions](#Other-Collection-Functions).

| What                                                         | How                                                          | Result                                                       |
| ------------------------------------------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------ |
| Create an empty collection:                                  | `Set dc = New DictCollection`                                |                                                              |
| Add an item with a key:                                      | `dc.Add "key1", 123.45`                                      | dc.Item(0) = 123.45                                          |
| Add an item without a key:                                   | `dc.Add , "A"` or `dc.Add "", "A"`                           | dc.Item(1) = "A"                                             |
| Get the number or items:                                     | `n = dc.Count`                                               | n = 2                                                        |
| Get the number or keys:                                      | `k = dc.KeyCount`                                            | k = 1                                                        |
| Get an item by key:                                          | `a = dc.Item("key1")`                                        | a = 123.45                                                   |
| Get an item by index:                                        | `a = dc.Item(0)`                                             | a = 123.45                                                   |
|                                                              | `b = dc.Item(1)`                                             | b = "A"                                                      |
| Default property `dc()` can be used as shorthand for `.Item()`: | `b = dc(1)`                                                  | b = "A"                                                      |
| Set an item by index:                                        | `dc(0) = 100`                                                | dc.Item(0) = 100                                             |
| Set an item by key:                                          | `dc("key1") = 200`                                           | dc.Item(0) = 200                                             |
| Add an item by key:                                          | `dc("key2") = "B"`                                           | dc.Item(2) = "B"                                             |
| Adding a nested DictCollection:                              | `dc.Add "key3", new DictCollection`                          | dc.Item(3) = [Empty DictCollection]                          |
| Add to nested DictCollection:                                | `dc("key3").Add "key4", 456.78`                              | dc.Item(3).Item(0) = 456.78                                  |
|                                                              | `dc("key3").Add , "C"`                                       | dc.Item(3).Item(1) = "C"                                     |
| Get items from nested DictCollection:                        | `c = dc("key3")("key4")`                                     | c = 456.78                                                   |
|                                                              | `d = dc("key3")(1)`                                          | d = "C"                                                      |
| Check if nested item exists:                                 | `dc("key3")(1) = dc.NonExistingValue`                        | false (yes it exists)                                        |
|                                                              | `dc("key3")("wrongkey") =   dc.NonExistingValue`             | true (no it doesn't)                                         |
| Get subcollection items even if they don't exist, also called 'endless subcollection chaining' | `a = dc("key3")("wrongkey")(4)("x")`                         | a = "[NONEXISTING]"                                          |
| Remove an item by key:                                       | `dc.Remove("key1")`                                          | "A" is now at index 0                                        |
| Remove an item by index:                                     | `dc.Remove(0)`                                               | "B" is now at index 0                                        |
| Get the index of an item by key:                             | `idx = dc.IndexOfKey("key3")`                                | idx = 1                                                      |
| Get an item's key (returns `Empty` if none exists):          | `a = dc.KeyOfItemAt(0)`                                      | a = "key2"                                                   |
| Check if key exists:                                         | `check = dc.Exists("key3")`                                  | check = True                                                 |
| Change an items key:                                         | `dc.Key("key2") = "key4"`                                    | "key2" changed to "key4"                                     |
| Change an items key while removing another item that used this key before (with `.ThrowErrors = False`): | `dc.Key("key4") = "key3"`                                    | "key4" changed to "key3", dc.Item(1)   removed               |
| Get keys for each item (include `Empty` values):             | `ka = dc.Keys`                                               | ka = Array("key3")                                           |
| Get all items in the added order:                            | `ia = dc.Items`                                              | ia = Array("B")                                              |
| Insert an item with key at index 0:                          | `dc.Insert "D", 0, "key5"`                                   | dc.Item(0) = "D", dc.Item(1) = "B"                           |
| Insert an item without key at index 3 (index -1 = at end):   | `dc.Insert "E", 3, ""`                                       | dc.Item(2) = Empty, dc.Item(3) = "E"                         |
| Get all items (include `Empty` values) as Array:             | `ia = dc.Items`                                              | ia = Array("D", "B", Empty,   "E")                           |
| Get all items without `Empty` values as Array:               | `ia = dc.Items(False)`                                       | ia = Array("D", "B", "E")                                    |
| Get all used keys in item order as Array:                    | `ka = dc.Keys(False)`                                        | ka(0) = Array("key5","key3")                                 |
| Get all keys in sorted order as Array:                       | `ka = dc.SortedKeys`                                         | ka = Array("key3","key5")                                    |
| Remove all items from DictCollection:                        | `dc.RemoveAll`                                               | All items are removed                                        |
| Move an item to another index:                               | `dc.Move(fromIndex, toIndex)`                                | Will expand storage to `toIndex` if `.ThrowErrors = False`   |
| Get an item by index (fast):                                 | `a = dc.ItemAt(Index)`                                       |                                                              |
| Set an item by index (fast):                                 | `dc.ItemAt(Index) = a`                                       |                                                              |
| Get an item by key (fast):                                   | `a = dc.ItemOf(Key)`                                         |                                                              |
| Set an item by key (fast):                                   | `dc.ItemOf(Key) = a`                                         |                                                              |
| Add (insert) a DictCollection:                               | `dc.AddDC(Key, AtIndex)`                                     | inserts a DictCollection at given index                      |
|                                                              | `dc.AddDC(Key)`                                              | adds DictCollection at the end                               |
|                                                              | `dc.AddDC(KeyA).AddDC(KeyB).AddDC(KeyC)`                     | adds multiple DictCollections after each other               |
| Get nested DictCollection:                                   | `dc.AsDC(KeyA).AsDC(KeyAA)`                                  | returns nested or nonexisting DictCollection                 |
| Get or add nested DictCollections:                           | `dc.SubDC(KeyA).SubDC(KeyAA)`                                | returns existing subkeys as DictCollections, nonexisting keys will be created as DictCollection |
| Chainable DictCollection functions: `.Add()`, `.Add2()`, `.AddDC()`, `.SetItem()`, `.Insert()`, `.Remove()`, `.RemoveAll()`, `.Move()` | `dc.RemoveAll().Add("a","A").Add("b","B")` `dc.Insert("D", -1).Insert("C", 2, "c")` | executes functions in chaining order: dc.Items = Array("A", "B", "C", "D")  `.Insert(…, -1, …)` inserts item at the end |

[Other Collection Functions](#other-collection-functions) are described after Settings.

### Settings

| Setting                                           | Explanation                                                  |
| ------------------------------------------------- | ------------------------------------------------------------ |
| `dc.ThowErrors = False` (default)                 | will return `.NonExistingValue` or `Empty` when accessing nonexistent indexes or keys |
| `dc.ThowErrors = True`                            | will throw errors when accessing nonexistent indexes or keys |
| `dc.LazySorting = True` (default)                 | will sort key array at first read access and thus speed up `.Add()` |
| `dc.LazySorting = False`                          | will sort key array at every `.Add()` if key is used         |
| `dc.CompareMode = 0` (default)                    | keys are case-sensitive, will binary compare keys, fastest   compare method |
| `dc.CompareMode = 1`                              | keys are case-insensitive, Ä = ä = A = a = á = Á             |
| `dc.CompareMode = 2`                              | keys are case-insensitive (only MSAccess), uses localeID for matching |
| `dc.EmulateDictionary = True`                     | DictCollection will behave like Scripting.Dictionary (no indexes, just keys) |
| `dc.EmulateCollection = True`                     | DictCollection will behave like VB Collection (`Collection.Add()` implemented as `DictCollection.Add2()`) |
| `dc.ZeroBasedIndex = True` (default)              | first item can be accessed with index = 0 (when ZeroBasedIndex=False, first index = 1) |
| `dc.DefaultValueEnabled = True` (default)         | calling default property without argument returns either the first item, `"[EMPTY]"`, or `"[NONEXISTING]"`, allows endless subcollection chaining: `Set dc2 = dc1("a")("nonexisting")("b")`   (if `.ThrowErrors = False`) |
| `dc.DefaultValueEnabled = False`                  | calling default property without argument as dc or dc()   returns a reference to the DictCollection, allows it to be used like a regular object: 'VarType(dc)' and   'dc Is DictCollection' will work |
| `dc.NonExistingValue = "[NONEXISTING]"` (default) | the value to be returned for nonexisting items if ThrowError=False; can be used to overwrite this value with anything except objects, e.g. `Empty`, `""`, `0`, `False` |
| `dc.EmptyCollectionValue = "[EMPTY]"` (default)   | the default value of empty DictCollections (ThrowError=False,   DefaultValueEnabled=True); can be used to overwrite this value with anything except   objects, e.g. Empty, "", 0, False |
| `dc.CollectionType`                               | sets/gets the collection type: 0 = NonExisting, 1 = Empty Array, 2 = Filled Array, 3 = Empty Key-Value-Store, 4 = Filled Key-Value-Store, 5 = Key-Value-Store with at least one item having no key |

### Other Collection Functions

| What                                                        | How                                                          |
| ----------------------------------------------------------- | ------------------------------------------------------------ |
| Collection-compatible Add function                          | `dc.Add2(Item, [Key], [Before], [After]) As DictCollection`  |
| Copy all items and keys                                     | `dc.CopyItems([SourceCollection], [TargetCollection],   [TargetIndex]) As DictCollection` |
| Copy settings from on DC to another DC                      | `dc.CopyAllSettingsExceptCollectionType(FromCollection,   ToCollection)` |
| Copy subcollection chaining settings                        | `dc.CopySubCollectionChainingSettings(FromCollection,   ToCollection)` |
| Clone entire DictCollection tree                            | `dc.Clone([TargetCollection]) as DictCollection`             |
| Set item by index or key (chainable function)               | `dc.SetItem(IndexOrKey, Value) As DictCollection`            |
| Increment stored number by 1 or Amount                      | `dc.Increment(IndexOrKey, [Amount]) As Variant`              |
| Check if item has key                                       | `dc.ItemHasKey(ItemIndex) As Boolean`                        |
| Check if item is a DictCollection                           | `dc.ItemIsDC(IndexOrKey) As Boolean`                         |
| Check if item is Object                                     | `dc.ItemIsObject(IndexOrKey) As Boolean`                     |
| Get item index associated with key (String)                 | `dc.IndexOfKey2(Key) As Long`                                |
| Get internal key storage index of an items key              | `dc.KeyIndexOfItemAt(Index) As Long`                         |
| Get key at internal key storage index                       | `dc.KeyAtKeyIndex(KeyIndex) As Variant`                      |
| Get string key at internal key storage index                | `dc.KeyAtKeyIndexAsString(KeyIndex) As Variant`              |
| Get an items string key from internal key storage           | `dc.KeyOfItemAtAsString(Index) As String`                    |
| Get all items and keys as Array(i,k)                        | `dc.ToArray() As Variant`                                    |
| Get items and keys as Array(i,k)                            | `dc.ItemsAndKeys([IncludeEmptyItems], [IncludeItemsWithoutKeys]) As Variant` |
| Get items and keys as Array(i,k) sorted by keys             | `dc.ItemsAndKeysSortedByKeys([IncludeEmptyItems],   [IncludeItemsWithoutKeys]) As Variant` |
| Get items as Array sorted by keys                           | `dc.ItemsSortedByKeys([IncludeEmptyItems], [IncludeItemsWithoutKeys]) As Variant` |
| Get keys and items as Array(k,i) sorted by keys             | d`c.SortedKeysAndItems([IncludeEmptyKeys], [IncludeEmptyItems]) As Variant` |
| Find key index where insert leaves keys sorted              | `dc.FindKeyInsertIndex(SearchedKey As Variant, [CompareMode]) As Long` |
| Find all keys that start with text and return as Array      | `dc.FindKeysThatStartWith(SearchText, [CompareMode]) As Variant` |
| Get infos about DictCollection tree content                 | `dc.AnalyzeDictCollectionTree([Recursive]) As DictCollection` with the following keys: |
|                                                             | `"NonStringConvertableItems"` - the number of the items in the subtree that are not convertable to String |
|                                                             | `"NonStringConvertableKeys"` - the number of the keys in the subtree that are not convertable to String |
|                                                             | `"SubCollections"` - the number of DictCollections in the subtree without counting items below circular references |
|                                                             | `"SubItems"` - the number of items in the subtree without counting items below circular references |
|                                                             | `"CircularReferences"` - the number of DictCollections in the subtree that contain DictCollections of upper levels of the same tree |
| Assign keys to items that have no keys in the format "_POS" | `dc.EnsureAllItemsHaveKeys()`                                |
| Convert DC tree to single key-value-store                   | `dc.Flatten([TargetCollection] As DictCollection,   [ReturnAnalysisResults]) As DictCollection` |
| Build DC tree from single key-value-store                   | `dc.Unflatten([TargetCollection])`                           |
| Demo basic functionality in Immediate Window                | `dc.DemoBasicFunctionality()`                                |
| Run all tests with output to Immediate Window               | `dc.SelfTest([DebugPrint])`                                  |
| Run compatibility tests and return errors                   | `dc.TestCompatibility([DebugPrint]) as Variant`              |
| Run functionality tests and return errors                   | `dc.TestFunctionality([DebugPrint]) as Variant`              |

### Utility Functions

| What                                                    | How                                                          |
| ------------------------------------------------------- | ------------------------------------------------------------ |
| Get missing argument (Error 448) as Variant             | `dc.UtilGetMissingValue([DoNotPassAnythingHere])`            |
| Assigns object or value to Variant variable             | `dc.UtilAssignFromTo(FromVariable, ToVariable)`              |
| Add value to array (creates one)                        | `dc.UtilAddArrayValue(Arr As Variant, Val As Variant)`       |
| Find index of value in array or return -1               | `dc.UtilFindArrayIndex(Arr, Val) As Long`                    |
| Remove a value from an array by index                   | `dc.UtilRemoveArrayValueByIndex(arr, Index)`                 |
| Get array dimensions (0 = uninitialized)                | `dc.UtilArrayDimensions(Arr) As Integer`                     |
| Remove a value from an array                            | `dc.UtilRemoveArrayValue(Arr, Val)`                          |
| Sort one/two-dimensional/nested array                   | `dc.UtilSortArray(Arr, FromIndex, ToIndex)`                  |
| Sort one/two-dimensional/nested array using `StrComp()` | `dc.UtilSortStringArray(Arr, FromIndex, ToIndex,   CompareMode)` |
| Check if text has only number chars                     | `dc.UtilStringConsistsOfNumericAsciiChars(Text) As   Boolean` |
| Build concatenated string by repeating a text           | `dc.UtilStringRepeat(Text, NumberOfTimes) As String`         |
| Check if text starts with another text                  | `dc.UtilStringStartsWith(Text, SearchText, [CompareMode]) As Boolean` |

### Demos and Use Cases (Todo)
#### Storing and retrieving data in object trees
#### Determining unique or distict keys
#### Getting a sorted file list from a folder/directory
#### Save DictCollection data to JSON
#### Migrate code that uses Scripting.Dictionary to DictCollection
#### Migrate code that uses Collection to DictCollection
#### Save Scripting.Dictionary data to JSON

## DictCollection.vbs

DictCollection.vbs is a stripped-down VBScript version of DictCollection.cls. It can be used to port VB6/VBA code that uses Collections to VBScript.

### Features

- **Versatile**: Can be used as Array, Key-Value-Store (Map) and Object Tree (JSON storage)
- **Compatible**: Can fully emulate Scripting.Dictionary or VB6/VBA Collection
- **Fast**: Retrieve items by index with array-like speed (faster than Dictionary and Collection)
- **Nonthrowing Design**: Does not throw errors by default (`dc.ThrowErrors=False`)
- **Endless Item Chaining**: `dc.Item("key1").Item("nonexisting_key").Item("key2")` always returns Collections that can be NONEXISTING, EMPTY or FILLED. Checking if a nested item exists is a one-liner.
- **Full Test Coverage**: SelfTest() covers all important functions and emulation features

### How To Use

| What                                                         | How                                             | Result                               |
| ------------------------------------------------------------ | ----------------------------------------------- | ------------------------------------ |
| Creating an empty collection:                                | `Set dc = New DictCollection`                   |                                      |
| Adding an item with a key:                                   | `dc.Add "key1", 123.45`                         | dc.Item(0) = 123.45                  |
| Adding an item without a key:                                | `dc.Add , "A"` or `dc.Add "", "A"`              | dc.Item(1) = "A"                     |
| Getting the number or items:                                 | `n = dc.Count`                                  | n = 2                                |
| Getting the number or keys:                                  | `k = dc.KeyCount`                               | k = 1                                |
| Getting an item by key:                                      | `a = dc.Item("key1")`                           | a = 123.45                           |
| Getting an item by index:                                    | `a = dc.Item(0)`                                | a = 123.45                           |
|                                                              | `b = dc.Item(1)`                                | b = "A"                              |
| Setting an item by index:                                    | `dc.Item(0) = 100`                              | dc.Item(0) = 100                     |
| Setting an item by key:                                      | `dc.Item("key1") = 200`                         | dc.Item(0) = 200                     |
| Adding an item by key:                                       | `dc.Item("key2") = "B"`                         | dc.Item(2) = "B"                     |
| Nesting another DictCollection:                              | `dc.Add "key3", new DictCollection`             | dc.Item(3) = [Empty DictCollection]  |
| Adding to nested DictCollection:                             | `dc.Item("key3").Add "key4", 456.78`            | dc.Item(3).Item(0) = 456.78          |
|                                                              | `dc.Item("key3").Add, "C"`                      | dc.Item(3).Item(1) = "C"             |
| Getting from nested DictCollection:                          | `c = dc.Item("key3").Item("key4")`              | c = 456.78                           |
|                                                              | `d = dc.Item("key3").Item(1)`                   | d = "C"                              |
| Checking if nested item exists:                              | `dc.Item("key3").Item(1) = dc.NonExistingValue` | False                                |
|                                                              | `dc.Item("key3").Item(2) = dc.NonExistingValue` | True                                 |
| Remove an item by key:                                       | `dc.Remove("key1")`                             | "A" is now at index=0                |
| Remove an item by index:                                     | `dc.Remove(0)`                                  | "B" is now at index=0                |
| Getting an item's key (returns `Empty` if none):             | `a = dc.KeyOfItemAt(0)`                         | a = "key2"                           |
| Checking if key exists:                                      | `check = dc.Exists("key3")`                     | check = True                         |
| Changing an items key:                                       | `dc.Key("key2") = "key4"`                       | "key2" changed to "key4"             |
| Changing an items key while removing another item that used this key before: | `dc.Key("key4") = "key3"`                       | "key4" "key3", dc.Item(1) removed    |
| Getting keys for each item incl. `Empty`:                    | `ka = dc.Keys`                                  | ka = Array("key3")                   |
| Getting all items in the added order:                        | `ia = dc.Items`                                 | ia = Array("B")                      |
| Inserting an item with key at index=0:                       | `dc.Insert "D", 0, "key5"`                      | dc.Item(0) = "D", dc.Item(1) = "B"   |
| Inserting an item without key at index=3:                    | `dc.Insert "E", 3, ""`                          | dc.Item(2) = Empty, dc.Item(3) = "E" |
| Getting all keys in sorted order:                            | `ka = dc.SortedKeys`                            | ka = Array("key3", "key5")           |
| Removing all items from DictCollection:                      | `dc.RemoveAll`                                  |                                      |

### Settings

| Setting                                             | Explanation                                                  |
| --------------------------------------------------- | ------------------------------------------------------------ |
| `dc.ThowErrors = False` (default)                   | will return `.NonExistingValue` or `Empty` when accessing nonexistent indexes or keys |
| `dc.ThowErrors = True`                              | will throw errors when accessing nonexistent indexes or keys |
| `dc.LazySorting = True` (default)                   | will sort key array at first access using a key and thus speed up `.Add` |
| `dc.LazySorting = False`                            | will sort key array at every `.Add` if key is used           |
| `dc.CompareMode = 0` (default)                      | keys are case-sensitive, will binary compare keys, fastest   compare method |
| `dc.CompareMode = 1`                                | keys are case-insensitive, Ä = ä = A = a = á = Á             |
| `dc.CompareMode = 2`                                | keys are case-insensitive (only MSAccess), uses localeID for matching |
| `dc.EmulateDictionary = True`                       | DictCollection will behave like Scripting.Dictionary (no indexes, just keys) |
| `dc.EmulateCollection = True`                       | DictCollection will behave like VB Collection (`Collection.Add` implemented as `DictCollection.Add2`) |
| `dc.ZeroBasedIndex = True` (default)                | first item can be accessed with index = 0 (when ZeroBasedIndex=False, first index = 1) |
| `dc.DefaultValueEnabled = True` (default)           | calling default property without argument returns either the   first item, "[EMPTY]", or "[NONEXISTING]", allows endless subcollection chaining: `Set dc2 = dc1("a")("nonexisting")("b")`   (ThrowErrors=False) |
| `dc.DefaultValueEnabled = False`                    | calling default property without argument as dc or dc()   returns a reference to the DictCollection, allows it to be used like a regular object: 'VarType(dc)' and   'dc Is DictCollection' will work |
| `dc.NonExistingValue = "[NONEXISTING]"` (default)   | the value to be returned for nonexisting items if ThrowError=False; can be used to overwrite this value with anything except objects, e.g. `Empty, "", 0, False` |
| `dc.EmptyCollectionValue = "[EMPTY]"` (default)     | the default value of empty DictCollections (ThrowError=False,   DefaultValueEnabled=True); can be used to overwrite this value with anything except objects, e.g. Empty, "", 0, False |
| `dc.CollectionType` sets/gets the collection type:  | 0 = NonExisting, 1 = Empty Array, 2 = Filled Array, 3 = Empty Key-Value-Store, 4 = Filled Key-Value-Store, 5 = Key-Value-Store with at least one item having no key |

### Other Collection Functions

| What                                                        | How                                                          |
| ----------------------------------------------------------- | ------------------------------------------------------------ |
| Collection-compatible Add function                          | `dc.Add2(Item, [Key], [Before], [After]) As DictCollection`  |
| Find all keys that start with text and return as Array      | `dc.FindKeysThatStartWith(SearchText, [CompareMode]) As Variant` |


### Utility Functions

| What                                                    | How                                                          |
| ------------------------------------------------------- | ------------------------------------------------------------ |
| Add value to array (creates one)                        | `dc.UtilAddArrayValue(Arr As Variant, Val As Variant)`       |
| Find index of value in array or return -1               | `dc.UtilFindArrayIndex(Arr, Val) As Long`                    |
| Remove a value from an array by index                   | `dc.UtilRemoveArrayValueByIndex(arr, Index)`                 |
| Get array dimensions (0 = uninitialized)                | `dc.UtilArrayDimensions(Arr) As Integer`                     |
| Remove a value from an array                            | `dc.UtilRemoveArrayValue(Arr, Val)`                          |
| Sort one/two-dimensional/nested array                   | `dc.UtilSortArray(Arr, FromIndex, ToIndex)`                  |
| Check if text starts with another text                  | `dc.UtilStringStartsWith(Text, SearchText, [CompareMode]) As Boolean` |


# Comparison of Collections in VB6 and VBA

## Available Options

1. **VB6/VBA Collection** object. Provided with the programming language this class allows adding items with or without key and retrieving them by key or by index. Keys can only be of datatype `String` and key matching is always case-insensitive. Collection supports iteration over its items with the `for each ... in ...` syntax and is optimized for fast adding and retrieving items by key. It can be created like this: `Set c = New Collection`. This object does not exist in VBScript.
2. **Scripting.Dictionary** comes with the "Microsoft Scripting Runtime" COM library (scrrun.dll) and is a very fast key-value-store that supports iteration over its items with the `for each ... in ...` syntax. Every item must have a key. Keys can be of any datatype or they can be objects. Number keys are data-type-insensitive which means that the key `1` as `Integer` = `1.0` as `Double` = `31/12/1899` as `Date`. String key matching can be either case-sensitive or case-insensitive. If the library is not referenced in your project (F2 > Right Mouse Click > References), you have to use the `Dim d as Object: Set d = CreateObject("Scripting.Dictionary")` syntax to create a new Dictionary object. Scripting.Dictionary does not exist on Apple Mac environments.
3. **DictCollection** is implemented in VB6/VBA and supports adding items with or without keys. It has two emulation modes `.EmulateDictionary = true` and `.EmulateCollection = true` that mimic the behavior of Dictionary and Collection. Keys have to be Strings and key matching can be case-sensitive or case-insensitive. Items and keys are stored internally as arrays so retrieving items by index is very fast. DictCollection has extended functionality like `.SortedKeys()`, `.ItemsAndKeys()`, `.Insert()` , `.Move()` and `.CopyItems()` and comes with useful String and Array functions like `.UtilAddArrayValue()`, `.UtilRemoveArrayValue()`, `.UtilRemoveArrayValueByIndex()` and `.UtilSortArray()`.

## Performance

1. **Retrieving items by index** is very fast in DictCollection, very slow in Collection (getting slower with every item added) and not supported in Dictionary (Workaround: use `.Items()` to copy all items to a new array and use that for retrieving items by index; this can be slow)
2. **Retrieving items by key** is very fast in Dictionary, fast in Collection and moderately fast in DictCollection
3. **Adding items without key** is very fast in Collection, moderately fast in DictCollection and not supported in Dictionary (Workaround: adding the index as key, very fast)
4. **Adding items with key** is very fast in Dictionary, fast in Collection and moderately fast in DictCollection

| Term            | Meaning                                                      |
| --------------- | ------------------------------------------------------------ |
| very fast       | processing 10'000 random items takes between 1ms and 10ms    |
| fast            | processing 10'000 random items takes between 10ms and 20ms   |
| moderately fast | processing 10'000 random items takes between 20ms and 80ms   |
| slow            | processing 10'000 random items takes between 80ms and 300ms  |
| very slow       | processing 10'000 random items takes between 300ms and 1000ms |

The performance of Scripting.Dictionary can be improved by adding a reference (F2 > Right Mouse Click > References) to "Microsoft Scripting Runtime" (scrrun.dll). This uses early binding instead of late binding and reduces the calling overhead for each operation.

## Functionality

|                                                              | Scripting.Dictionary                               | Collection                                         | DictCollection                                               |
| ------------------------------------------------------------ | -------------------------------------------------- | -------------------------------------------------- | ------------------------------------------------------------ |
|                                                              | or DictCollection with `.EmulateDictionary = True` | or DictCollection with `.EmulateCollection = True` |                                                              |
| Has Dictionary style `.Add(Key, Item)` function              | yes                                                | no                                                 | yes, `.Add()`                                                |
| Has Collection style `.Add(Item, Key, Before, After)` function | no                                                 | yes                                                | yes, `.Add2()`                                               |
| Has `.Insert()` function                                     | no                                                 | no                                                 | yes                                                          |
| Supports Iteration using `For ... Each ...   Next`           | yes                                                | yes                                                | no                                                           |
| Zero-based index                                             | yes                                                | no, 1-based                                        | configurable `.ZeroBasedIndex = True`                        |
| Items without Keys allowed                                   | no                                                 | yes                                                | yes                                                          |
| String `""` is allowed as key                                | yes                                                | yes                                                | no, same as no key                                           |
| VBA Keyword Empty is allowed as key                          | yes                                                | yes                                                | no, same as no key                                           |
| Numbers/Dates allowed as Key                                 | yes                                                | no                                                 | no                                                           |
| Objects allowed as key                                       | yes                                                | no                                                 | no                                                           |
| Missing (or omitted argument) allowed as key                 | yes                                                | no, same as no key                                 | no, same as no key                                           |
| Arrays allowed as key                                        | no                                                 | no                                                 | no                                                           |
| Case insensitive key matching ("A" = "a")                    | configurable `.CompareMode = 1`                    | yes                                                | configurable `.CompareMode = 1`                              |
| Case sensitive key matching ("A" <> "a")                     | configurable `.CompareMode = 0`                    | no                                                 | configurable `.CompareMode = 0`                              |
| Empty, Missing, Null, Nothing, Errors allowed as items       | yes                                                | yes                                                | yes                                                          |
| Numeric data types and Strings allowed as items              | yes                                                | yes                                                | yes                                                          |
| Objects, Arrays and Variants allowed as items                | yes                                                | yes                                                | yes                                                          |
| Implicit add using `.Item()` with nonexisting keys           | yes                                                | no                                                 | no                                                           |
| Throws error when accessing nonexisting keys                 | yes                                                | no                                                 | configurable `.ThrowErrors = true`                           |
| Throws error when re-adding existing keys                    | yes                                                | yes                                                | configurable `.ThrowErrors = true`                           |
| Throws error when removing nonextisting keys                 | yes                                                | yes                                                | configurable `.ThrowErrors = true`                           |
| Throws error when changing keys to used keys                 | yes                                                | -                                                  | configurable `.ThrowErrors = true`                           |
| Returns Empty when accessing nonexisting keys                | yes                                                | no                                                 | configurable `.NonExistingValue = Empty` + `.ThrowErrors = False` |
| Has `.Count()` function that returns number of items         | yes                                                | yes                                                | yes                                                          |
| Has `.KeyCount()` function that returns number of keys       | yes, same as `.Count()`                            | no                                                 | yes                                                          |
| Has `.Exists()` function                                     | yes                                                | no                                                 | yes                                                          |
| Has `.Items()` function                                      | yes                                                | no                                                 | yes                                                          |
| `.Items()` can exclude all Empty values                      | no                                                 | no                                                 | yes, no in VBScript                                          |
| Has `.Keys()` function                                       | yes                                                | no                                                 | yes                                                          |
| `.Keys()` can exclude all Empty values                       | no                                                 | -                                                  | yes, no in VBScript                                          |
| Has `.Key()` function to change keys                         | yes                                                | no                                                 | yes                                                          |
| Has `.KeyOfItemAt()` function                                | no                                                 | no                                                 | yes                                                          |
| Has `.IndexOfKey()`                                          | no                                                 | no                                                 | yes                                                          |
| `.Item(IndexOrKey)` is default property                      | no                                                 | no                                                 | yes, no in VBScript                                          |
| Allows endless subcollection chaining, e.g. `Set dc2 = dc1("a")("nonexisting")("b")` | no                                                 | no                                                 | configurable `.DefaultValueEnabled = True` + `ThrowErrors = False` |
| Default property returns first item as default value in filled collections | no                                                 | no                                                 | configurable `.DefaultValueEnabled = true`                   |