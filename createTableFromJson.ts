function main(workbook: ExcelScript.Workbook) {
    // Delete the "Combined" worksheet, if it's present.
    workbook.getWorksheet("Combined")?.delete();

    // Create a new worksheet named "Combined" for the combined table.
    const newSheet = workbook.addWorksheet("Combined");

    // Get the header values for the first table in the workbook.
    // This also saves the table list before we add the new, combined table.
    const tables = workbook.getTables();
    const headerValues = tables[0].getHeaderRowRange().getTexts();
    console.log(headerValues);

    // Copy the headers on a new worksheet to an equal-sized range.
    const targetRange = newSheet
        .getRange("A1")
        .getResizedRange(headerValues.length - 1, headerValues[0].length - 1);
    targetRange.setValues(headerValues);

    // Add the data from each table in the workbook to the new table.
    const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
    for (let table of tables) {
        // Get all the values from the table as text.
        let texts = table.getRange().getTexts();
        let rowCount = table.getRowCount();

        // Create an array of JSON objects that match the row structure.
        let returnObjects: TableData[] = [];
        if (rowCount > 0) {
            returnObjects = returnObjectFromValues(texts);
        }

        let levenshteinDistance = (a: string, b: string): number => {
            if (a.length === 0) return b.length;
            if (b.length === 0) return a.length;

            let matrix: number[][] = [];

            for (let i = 0; i <= b.length; i++) {
                matrix[i] = [i];
            }

            for (let j = 0; j <= a.length; j++) {
                matrix[0][j] = j;
            }

            for (let i = 1; i <= b.length; i++) {
                for (let j = 1; j <= a.length; j++) {
                    if (b.charAt(i - 1) === a.charAt(j - 1)) {
                        matrix[i][j] = matrix[i - 1][j - 1];
                    } else {
                        matrix[i][j] = Math.min(
                            matrix[i - 1][j - 1] + 1,
                            matrix[i][j - 1] + 1,
                            matrix[i - 1][j] + 1
                        );
                    }
                }
            }

            return matrix[b.length][a.length];
        };

        let result = returnObjects.reduce((acc: TableData[], cur) => {
            let existing = acc.find((x) => {
                let curName = cur.Name.toLowerCase()
                    .replace(/\(.*\)/, "")
                    .replace(/ +$/, "");
                let xName = x.Name.toLowerCase()
                    .replace(/\(.*\)/, "")
                    .replace(/ +$/, "");
                return curName === xName;
            });
            if (existing) {
                existing.Count += 1;
                existing.Value += parseInt(cur.Value);
            } else {
                acc.push({
                    Name: cur.Name,
                    Count: 1,
                    Value: parseInt(cur.Value),
                });
            }
            return acc;
        }, []);

        let values = result.map((x) => [x.Name, x.Count, x.Value]);

        // If the table is not empty, add its rows to the combined table.
        if (rowCount > 0) {
            combinedTable.addRows(-1, values);
        }
    }

    return "OK";
}

// This function converts a 2D array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
    let objectArray: TableData[] = [];
    let objectKeys: string[] = [];
    for (let i = 0; i < values.length; i++) {
        if (i === 0) {
            objectKeys = values[i];
            continue;
        }

        let object = {};
        for (let j = 0; j < values[i].length; j++) {
            object[objectKeys[j]] = values[i][j];
        }

        objectArray.push(object as TableData);
    }

    return objectArray;
}

interface TableData {
    Name: string;
    Count: number;
    Value: number;
}
