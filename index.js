const ExcelJS = require('exceljs');

async function createExcelWithTable() {
    // Crear un nuevo libro de trabajo
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Ventas 2024');

    // Definir los datos
    const data = [
        { id: 1, producto: 'Laptop', precio: 999.99, stock: 50, categoria: 'Electrónica' },
        { id: 2, producto: 'Monitor', precio: 299.99, stock: 30, categoria: 'Electrónica' },
        { id: 3, producto: 'Teclado', precio: 79.99, stock: 100, categoria: 'Accesorios' },
        { id: 4, producto: 'Mouse', precio: 29.99, stock: 150, categoria: 'Accesorios' },
        { id: 5, producto: 'Webcam', precio: 59.99, stock: 45, categoria: 'Periféricos' }
    ];

    // Crear una tabla
    worksheet.addTable({
        name: 'TablaProductos',
        ref: 'A1',
        headerRow: true,
        totalsRow: true,
        style: {
            theme: 'TableStyleMedium2',
            showRowStripes: true,
        },
        columns: [
            { name: 'ID', totalsRowLabel: 'Total:', filterButton: true },
            { name: 'Producto', filterButton: true },
            { name: 'Precio', totalsRowFunction: 'average', filterButton: true },
            { name: 'Stock', totalsRowFunction: 'sum', filterButton: true },
            { name: 'Categoría', filterButton: true }
        ],
        rows: data.map(item => [
            item.id,
            item.producto,
            item.precio,
            item.stock,
            item.categoria
        ])
    });

    // Ajustar el ancho de las columnas
    worksheet.columns.forEach(column => {
        column.width = 15;
    });

    // Guardar el archivo
    try {
        await workbook.xlsx.writeFile('inventario.xlsx');
        console.log('Archivo Excel creado exitosamente!');
    } catch (error) {
        console.error('Error al crear el archivo:', error);
    }
}

// Ejecutar la función
createExcelWithTable();