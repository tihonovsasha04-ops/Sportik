const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const cors = require('cors');
const bodyParser = require('body-parser');
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');


const app = express();
const PORT = 3000;
const db = new sqlite3.Database('./database.db');

app.use(cors());
app.use(bodyParser.json());
app.use('/images', express.static(path.join(__dirname, 'images')));
app.use(express.static('public'));

app.listen(PORT, () => {
    console.log(`Сервер працює на http://localhost:${PORT}`);
});

const storage = multer.diskStorage({
    destination: './images/',
    filename: (req, file, cb) => {
        cb(null, Date.now() + path.extname(file.originalname));
    }
});
const upload = multer({ storage });

// Створення таблиці, якщо вона не існує
const createTableQuery = `CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    image TEXT,
    material TEXT,
    size TEXT,
    description TEXT,
    manufacturer TEXT,
    quantity INTEGER,
    price REAL,
    total_price REAL,
    delivery_date TEXT,
    supplier TEXT,
    availability TEXT
)`;
db.run(createTableQuery);

// Додавання товару
app.post('/products', upload.single('image'), (req, res) => {
    const { name, material, size, description, manufacturer, quantity, price, delivery_date, supplier, availability } = req.body;
    if (!name) return res.status(400).json({ error: "Назва товару є обов'язковою" });

    const image = req.file ? `/images/${req.file.filename}` : null;
    const total_price = Number(quantity) * Number(price);

    const query = `INSERT INTO products (name, image, material, size, description, manufacturer, quantity, price, total_price, delivery_date, supplier, availability) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;
    db.run(query, [name, image, material, size, description, manufacturer, quantity, price, total_price, delivery_date, supplier, availability], function (err) {
        if (err) return res.status(500).json({ error: err.message });
        res.json({ id: this.lastID });
    });
});

app.put('/products/:id', upload.single('image'), (req, res) => {
    const { name, material, size, description, manufacturer, quantity, price, delivery_date, supplier, availability } = req.body;
    if (!name) return res.status(400).json({ error: "Назва товару є обов'язковою" });

    const image = req.file ? `/images/${req.file.filename}` : null;
    const total_price = Number(quantity) * Number(price);

    const query = `UPDATE products SET name = ?, image = COALESCE(?, image), material = ?, size = ?, description = ?, manufacturer = ?, quantity = ?, price = ?, total_price = ?, delivery_date = ?, supplier = ?, availability = ? WHERE id = ?`;
    db.run(query, [name, image, material, size, description, manufacturer, quantity, price, total_price, delivery_date, supplier, availability, req.params.id], function (err) {
        if (err) return res.status(500).json({ error: err.message });
        res.json({ updated: this.changes });
    });
});

// Отримання всіх товарів або пошук за назвою
app.get('/products', (req, res) => {
    const { search, minPrice, maxPrice, manufacturer, availability } = req.query;
    let query = 'SELECT * FROM products WHERE 1=1';
    let params = [];

    if (search) {
        query += ' AND name LIKE ?';
        params.push(`%${search}%`);
    }
    if (minPrice) {
        query += ' AND price >= ?';
        params.push(minPrice);
    }
    if (maxPrice) {
        query += ' AND price <= ?';
        params.push(maxPrice);
    }
    if (manufacturer) {
        query += ' AND manufacturer = ?';
        params.push(manufacturer);
    }
    if (availability) {
        query += ' AND availability = ?';
        params.push(availability);
    }

    db.all(query, params, (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(rows);
    });
});

// Оновлення товару
app.put('/products/:id', (req, res) => {
    const { name, material, size, description, manufacturer, quantity, price, delivery_date, supplier } = req.body;
    const total_price = quantity * price;
    db.run(`UPDATE products SET name = ?, material = ?, size = ?, description = ?, manufacturer = ?, quantity = ?, price = ?, total_price = ?, delivery_date = ?, supplier = ? WHERE id = ?`,
        [name, material, size, description, manufacturer, quantity, price, total_price, delivery_date, supplier, req.params.id],
        function (err) {
            if (err) return res.status(500).json({ error: err.message });
            res.json({ updated: this.changes });
        }
    );
});

// Видалення товару
app.delete('/products/:id', (req, res) => {
    db.run(`DELETE FROM products WHERE id = ?`, req.params.id, function (err) {
        if (err) return res.status(500).json({ error: err.message });
        res.json({ deleted: this.changes });
    });
});

app.get('/manufacturers', (req, res) => {
    db.all('SELECT DISTINCT manufacturer FROM products', [], (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(rows.map(row => row.manufacturer));
    });
});

app.get('/supply-data', (req, res) => {
    const { startDate, endDate } = req.query;
    if (!startDate || !endDate) {
        return res.status(400).json({ error: "Потрібно вказати діапазон дат" });
    }

    const query = `
        SELECT delivery_date, SUM(quantity) as total_quantity
        FROM products
        WHERE delivery_date BETWEEN ? AND ?
        AND availability = 'Є в наявності'
        GROUP BY delivery_date
        ORDER BY delivery_date
    `;

    db.all(query, [startDate, endDate], (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(rows);
    });
});

// === ЕКСПОРТ ТОВАРІВ У EXCEL ===
app.get('/export/excel', (req, res) => {
    db.all('SELECT * FROM products', [], (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });

        // створюємо аркуш і книгу
        const worksheet = xlsx.utils.json_to_sheet(rows);
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Товари');

        // зберігаємо у файл
        const filePath = path.join(__dirname, 'products_export.xlsx');
        xlsx.writeFile(workbook, filePath);

        // відправляємо клієнту
        res.download(filePath, 'products.xlsx');
    });
});

// === ІМПОРТ ТОВАРІВ З EXCEL ===
app.post('/import/excel', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'Файл не завантажено' });

    const filePath = req.file.path;
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const insert = db.prepare(`
        INSERT INTO products (name, image, material, size, description, manufacturer, quantity, price, total_price, delivery_date, supplier, availability)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    data.forEach(p => {
        insert.run(
            p.name,
            p.image || '',
            p.material || '',
            p.size || '',
            p.description || '',
            p.manufacturer || '',
            p.quantity || 0,
            p.price || 0,
            (p.quantity || 0) * (p.price || 0),
            p.delivery_date || '',
            p.supplier || '',
            p.availability || ''
        );
    });

    insert.finalize();
    res.json({ imported: data.length });
});




