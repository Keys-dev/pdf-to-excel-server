import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, HeadingLevel } from "docx";
function buildTable(table) {
    const columnCount = Math.max(...table.rows.map((row) => row.length), 1);
    const rows = table.rows.map((row) => new TableRow({
        children: row.map((cell) => new TableCell({
            width: { size: 100 / columnCount, type: WidthType.PERCENTAGE },
            children: [new Paragraph(cell)],
        })),
    }));
    return new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
    });
}
export async function tablesToWordBuffer(tables, fileName) {
    const children = [
        new Paragraph({ text: fileName, heading: HeadingLevel.HEADING_1 }),
    ];
    if (tables.length === 0) {
        children.push(new Paragraph({ text: "No tables were detected in this PDF." }));
    }
    tables.forEach((table, index) => {
        children.push(new Paragraph({ text: `Table ${index + 1} (page ${table.page})`, heading: HeadingLevel.HEADING_2 }));
        children.push(buildTable(table));
        children.push(new Paragraph({ text: "" }));
    });
    const doc = new Document({
        sections: [{ children }],
    });
    return Packer.toBuffer(doc);
}
