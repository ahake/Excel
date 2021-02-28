# How to get the last cell in table column:

Use COUNTA to get count of rows:
```basic
=COUNTA(Table[Column])
```

Then use that value to offset:
```basic
=OFFSET(TheCellOneAboveTheTable,COUNTA(Table[Column],0)
```

# Explanation

|Value|Description|
|:---|:---|
|TheCellOneAboveTheTable|We need to count the rows of the entire table, but if we offset the first row of the table with all the rows, we end up one below the table. Therefore we need to start with the cell just above the column|
|Table|The table containing the column|
|Column|The column with the cell we're after|
|COUNTA|Counts cells until it finds an empty cell, so make sure there are no empty cells in the table!|
|OFFSET|Returns the value of the cell in the offset|