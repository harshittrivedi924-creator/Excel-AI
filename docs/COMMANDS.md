# Command Reference

## Aggregate Operations

| Command | Example | Description |
|---------|---------|-------------|
| SUM | `Column D ka sum karo` | Sum all values in column |
| AVERAGE | `Revenue ka average nikal do` | Average of column values |
| MIN | `Sales ka minimum batao` | Find minimum value |
| MAX | `Column B ka maximum` | Find maximum value |
| COUNT | `Price ka count karo` | Count numeric values |

### With Destination

```
Column D ka total karo aur D21 mein daal do
Revenue ka average nikal ke F10 mein likh do
```

### Named Columns

```
Sales ka total karo
Expense ka average karo
Profit ka sum nikalo
```

### Whole Sheet

```
Sabka sum kar do
All ka average nikalo
Poora sheet ka total
```

## Arithmetic Operations

| Command | Example | Description |
|---------|---------|-------------|
| ADD | `B2 + C2` | Add two cells |
| SUBTRACT | `B2 minus C2` | Subtract values |
| MULTIPLY | `Revenue guna Price` | Multiply columns |
| DIVIDE | `B2 div C2` | Divide values |

### Element-wise (Named Columns)

```
Revenue minus expense karke profit nikalo
B aur C ko multiply karke D mein daal do
Sales plus Returns karo
```

### Cell References

```
B2 + 100
C3 * D3
A1 minus B1
```

## Advanced Operations

| Command | Example | Description |
|---------|---------|-------------|
| DIFFERENCE | `B2 aur C2 ka difference` | Absolute difference |
| GROWTH | `February vs January kitni badhi` | % growth between values |
| PERCENTAGE | `Sales ka percentage karma karo` | % of total |

### Growth Examples

```
February ki sales January se kitni badhi
B2 se B3 kitni badh gayi
```

### Percentage Examples

```
Revenue ka percentage nikalo
Sales ka percent calculate karo
```

## Conditional Operations

| Command | Example | Description |
|---------|---------|-------------|
| SUMIF | `Column B ka sum jahan A Pen hai` | Conditional sum |
| COUNTIF | `Count karo jahan B 100 se zyada` | Conditional count |
| AVERAGEIF | `Revenue ka average jahan Month March hai` | Conditional average |

### Conditions

```
# Numeric comparisons
jahan Price 100 se zyada
jahan Revenue > 1000
jahan Cost <= 500

# Text equality
jahan Product Pen hai
jahan Category Electronics

# Multiple conditions
B ka sum jahan A Pen hai aur C 50 se kam
Revenue ka average jahan Month March hai aur Status Active
```

## Data Operations

| Command | Example | Description |
|---------|---------|-------------|
| SORT | `Column B sort karo` | Sort by column |
| DEDUPE | `Duplicate rows hata do` | Remove duplicates |
| FILTER | `Filter data jahan Revenue 100 se zyada` | Filter rows |
| FIND_EMPTY | `Empty cells batao` | Find empty cells |

### Sort Examples

```
Column D sort karo
B arrange karo
```

### Filter Examples

```
Filter data jahan Revenue 1000 se zyada
Filter karo jahan Status Active
Revenue > 500 wale rows dikhao
```

## Data Cleaning

| Command | Example | Description |
|---------|---------|-------------|
| STANDARDIZE_DATES | `Dates ko standardize karo` | Format dates to DD-MM-YYYY |
| STANDARDIZE_NAMES | `Names ko title case karo` | Title-case text cells |
| DETECT_INVALID | `Invalid values batao` | Find data issues |

## Visualization

| Command | Example | Description |
|---------|---------|-------------|
| CHART | `Sales ka chart bana do` | Create bar chart |
| CHART (line) | `Revenue ka line chart bana do` | Create line chart |

## Analysis

| Command | Example | Description |
|---------|---------|-------------|
| ANALYZE | `Is workbook ka analysis kar do` | Full workbook analysis |

Analysis includes:
- Column metrics (total, average, min, max, count)
- Trend detection (biggest increases/decreases)
- Anomaly detection (values > 2 std devs from mean)

## Hinglish Keywords

| English | Hinglish |
|---------|----------|
| sum, total | sigma, jod |
| average | avg, mean |
| minimum | min |
| maximum | max |
| add | plus, jod |
| subtract | minus, ghatao |
| multiply | times, product, guna |
| divide | div, bhaag |
| difference | antar |
| growth | increase, badhi, kitni badhi |
| percentage | percent, kitna percent |
| sort | arrange |
| duplicates | remove duplicates, duplicate rows hata |
| filter | filter karo |
| empty | khali |
| chart | graph, chart bana |
| analyze | analyse, report bana |
| where | jahan |
| and | aur |
