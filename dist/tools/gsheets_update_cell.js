import { google } from "googleapis";
export const schema = {
    name: "gsheets_update_cell",
    description: "Update a cell value in a Google Spreadsheet",
    inputSchema: {
        type: "object",
        properties: {
            fileId: {
                type: "string",
                description: "ID of the spreadsheet",
            },
            range: {
                type: "string",
                description: "Cell range in A1 notation (e.g. 'Sheet1!A1')",
            },
            value: {
                type: "string",
                description: "New cell value",
            },
        },
        required: ["fileId", "range", "value"],
    },
};
export async function updateCell(args) {
    const { fileId, range, value } = args;
    const sheets = google.sheets({ version: "v4" });
    const isFormula = value.toString().startsWith("=");
    await sheets.spreadsheets.values.update({
        spreadsheetId: fileId,
        range: range,
        valueInputOption: isFormula ? "USER_ENTERED" : "RAW",
        requestBody: {
            values: [[value]],
        },
    });
    return {
        content: [
            {
                type: "text",
                text: `Updated cell ${range} to value: ${value}`,
            },
        ],
        isError: false,
    };
}
export const batchSchema = {
    name: "gsheets_update_cells_batch",
    description: "Update multiple cells in a Google Spreadsheet in a single batch operation",
    inputSchema: {
        type: "object",
        properties: {
            fileId: {
                type: "string",
                description: "ID of the spreadsheet",
            },
            updates: {
                type: "array",
                description: "Array of cell updates to perform",
                items: {
                    type: "object",
                    properties: {
                        range: {
                            type: "string",
                            description: "Cell range in A1 notation (e.g. 'Sheet1!A1')",
                        },
                        value: {
                            type: "string",
                            description: "New cell value",
                        },
                    },
                    required: ["range", "value"],
                },
            },
        },
        required: ["fileId", "updates"],
    },
};
export async function updateCellsBatch(args) {
    const { fileId, updates } = args;
    const sheets = google.sheets({ version: "v4" });
    // Prepare batch update data
    const data = updates.map((update) => ({
        range: update.range,
        values: [[update.value]],
    }));
    // Check if any values are formulas to determine value input option
    const hasFormulas = updates.some((update) => update.value.toString().startsWith("="));
    await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: fileId,
        requestBody: {
            valueInputOption: hasFormulas ? "USER_ENTERED" : "RAW",
            data: data,
        },
    });
    const updatedRanges = updates.map(update => `${update.range}: ${update.value}`).join(", ");
    return {
        content: [
            {
                type: "text",
                text: `Batch updated ${updates.length} cells: ${updatedRanges}`,
            },
        ],
        isError: false,
    };
}
