# JSON Format for AI Agent Responses

This document outlines the expected JSON structure that the LLM (Language Model) should output to be processed by the Excel AI Agent services and protocol.

## Basic Response Structure

```json
{
  "message": "A brief message to display to the user",
  "body": "Optional formatted content to display in the chat (e.g., tables, code blocks)",
  "actions": [
    {
      "type": "ACTION_TYPE",
      "description": "Optional description of what this action does",
      "data": {
        // Action-specific data
      }
    }
  ],
  "metadata": {
    "success": true,
    "errors": [],
    "additionalData": "Any additional information"
  }
}
```

## Action Types

The `type` field in each action should be one of the following values:

- `WRITE_CELL`: Write a value to a cell
- `READ_CELL`: Read a value from a cell
- `FORMAT_CELL`: Format a cell
- `WRITE_RANGE`: Write values to a range
- `READ_RANGE`: Read values from a range
- `FORMAT_RANGE`: Format a range
- `CREATE_WORKSHEET`: Create a new worksheet
- `DELETE_WORKSHEET`: Delete a worksheet
- `RENAME_WORKSHEET`: Rename a worksheet
- `INSERT_FORMULA`: Insert a formula
- `CREATE_CHART`: Create a chart
- `CREATE_PIVOT_TABLE`: Create a pivot table
- `APPLY_FILTER`: Apply a filter
- `APPLY_CONDITIONAL_FORMATTING`: Apply conditional formatting
- `APPLY_DATA_VALIDATION`: Apply data validation
- `CUSTOM`: Custom action

## Action Data Examples

### WRITE_CELL

```json
{
  "type": "WRITE_CELL",
  "description": "Write the value 42 to cell A1",
  "data": {
    "address": "A1",
    "value": 42,
    "formatting": {
      "font": {
        "bold": true,
        "color": "#FF0000"
      }
    }
  }
}
```

### WRITE_RANGE

```json
{
  "type": "WRITE_RANGE",
  "description": "Write values to range A1:B3",
  "data": {
    "address": "A1:B3",
    "values": [
      [1, 2],
      [3, 4],
      [5, 6]
    ],
    "formatting": {
      "border": {
        "style": "thin",
        "color": "#000000"
      }
    }
  }
}
```

### READ_RANGE

```json
{
  "type": "READ_RANGE",
  "description": "Read values from range A1:Z100",
  "data": {
    "address": "A1:Z100"
  }
}
```

### CREATE_PIVOT_TABLE

```json
{
  "type": "CREATE_PIVOT_TABLE",
  "description": "Create a pivot table from the data in the current worksheet",
  "data": {
    "sourceRange": "A1:D10",
    "destinationRange": "F1",
    "rows": ["Category"],
    "columns": ["Region"],
    "values": [
      {
        "field": "Sales",
        "function": "sum"
      }
    ]
  }
}
```

### CREATE_CHART

```json
{
  "type": "CREATE_CHART",
  "description": "Create a column chart from the data in the current worksheet",
  "data": {
    "type": "column",
    "title": "Sales by Region",
    "dataRange": "A1:B5",
    "destinationRange": "D1:H10"
  }
}
```

### INSERT_FORMULA

```json
{
  "type": "INSERT_FORMULA",
  "description": "Insert a SUM formula to calculate the total sales",
  "data": {
    "address": "B10",
    "formula": "=SUM(B2:B9)"
  }
}
```

## Complete Response Example

```json
{
  "message": "I've analyzed your data and created a summary in cell A1.",
  "body": "Here's a summary of the data:\n\n| Category | Sales |\n|----------|-------|\n| Product A | 100 |\n| Product B | 200 |\n| Product C | 300 |",
  "actions": [
    {
      "type": "WRITE_CELL",
      "description": "Write the summary title to cell A1",
      "data": {
        "address": "A1",
        "value": "Sales Summary",
        "formatting": {
          "font": {
            "bold": true,
            "size": 14
          }
        }
      }
    },
    {
      "type": "WRITE_RANGE",
      "description": "Write the summary data to range A2:B5",
      "data": {
        "address": "A2:B5",
        "values": [
          ["Category", "Sales"],
          ["Product A", 100],
          ["Product B", 200],
          ["Product C", 300]
        ],
        "formatting": {
          "border": {
            "style": "thin",
            "color": "#000000"
          }
        }
      }
    }
  ],
  "metadata": {
    "success": true,
    "timestamp": "2023-06-15T12:34:56Z"
  }
}
```

## Guidelines for LLM Output

1. Always include a `message` field with a brief explanation of what the response does.
2. Use the `body` field for formatted content that should be displayed in the chat.
3. Include one or more `actions` that specify what operations to perform in Excel.
4. Each action must have a `type` and `data` field.
5. The `data` field should contain all the information needed to perform the action.
6. Optionally include a `description` field to explain what the action does.
7. Optionally include `metadata` with additional information about the response.

By following this structure, the LLM's output can be directly processed by the Excel AI Agent services and protocol. 