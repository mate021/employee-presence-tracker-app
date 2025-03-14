const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const multer = require('multer');
const path = require('path');
const XLSX = require('xlsx');
const fs = require('fs');

// Initialize express app
const app = express();
const port = process.env.PORT || 3000;

// Set up static file serving
app.use(express.static(__dirname));
app.use(express.json());

// Set up file upload storage
const storage = multer.diskStorage({
    destination: function(req, file, cb) {
        const uploadsDir = path.join(__dirname, 'uploads');
        if (!fs.existsSync(uploadsDir)) {
            fs.mkdirSync(uploadsDir);
        }
        cb(null, uploadsDir);
    },
    filename: function(req, file, cb) {
        cb(null, file.fieldname + '-' + Date.now() + path.extname(file.originalname));
    }
});

const upload = multer({ 
    storage: storage,
    fileFilter: function(req, file, cb) {
        if (path.extname(file.originalname) !== '.xlsx') {
            return cb(new Error('Only .xlsx files are allowed'));
        }
        cb(null, true);
    }
});

// Initialize SQLite database
const db = new sqlite3.Database('./employee_presence.db', (err) => {
    if (err) {
        console.error('Error opening database', err);
    } else {
        console.log('Connected to the SQLite database');
        createTables();
    }
});

// Create database tables
function createTables() {
    db.serialize(() => {
        // Employees table
        db.run(`CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT UNIQUE,
            name TEXT,
            directorate TEXT,
            department TEXT
        )`);

        // Card logins table
        db.run(`CREATE TABLE IF NOT EXISTS card_logins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT,
            date TEXT,
            time TEXT,
            FOREIGN KEY (employee_id) REFERENCES employees (employee_id)
        )`);

        // Official leaves table
        db.run(`CREATE TABLE IF NOT EXISTS official_leaves (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id TEXT,
            date TEXT,
            reason TEXT,
            FOREIGN KEY (employee_id) REFERENCES employees (employee_id)
        )`);
    });
}

// API endpoint to upload employees data
app.post('/api/upload/employees', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
    }

    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        // Clear existing data
        db.run('DELETE FROM employees', [], (err) => {
            if (err) {
                return res.status(500).json({ error: 'Failed to clear existing data' });
            }

            // Insert new data
            const stmt = db.prepare('INSERT INTO employees (employee_id, name, directorate, department) VALUES (?, ?, ?, ?)');
            
            data.forEach(row => {
                const employeeId = row.id || row.employeeId || row.employee_id;
                const name = row.name || row.employeeName || row.employee_name || '';
                const directorate = row.directorate || row.directoriate || '';
                const department = row.department || '';
                
                stmt.run(employeeId, name, directorate, department);
            });
            
            stmt.finalize();
            res.json({ success: true, count: data.length });
        });
    } catch (error) {
        console.error('Error processing employees file:', error);
        res.status(500).json({ error: 'Failed to process file' });
    }
});

// API endpoint to upload card login data
app.post('/api/upload/card-login', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
    }

    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        // Clear existing data
        db.run('DELETE FROM card_logins', [], (err) => {
            if (err) {
                return res.status(500).json({ error: 'Failed to clear existing data' });
            }

            // Insert new data
            const stmt = db.prepare('INSERT INTO card_logins (employee_id, date, time) VALUES (?, ?, ?)');
            
            data.forEach(row => {
                const employeeId = row.employeeId || row.employee_id || row.id;
                const date = row.date || '';
                const time = row.time || '';
                
                stmt.run(employeeId, date, time);
            });
            
            stmt.finalize();
            res.json({ success: true, count: data.length });
        });
    } catch (error) {
        console.error('Error processing card login file:', error);
        res.status(500).json({ error: 'Failed to process file' });
    }
});

// API endpoint to upload official leaves data
app.post('/api/upload/official-leaves', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
    }

    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        // Clear existing data
        db.run('DELETE FROM official_leaves', [], (err) => {
            if (err) {
                return res.status(500).json({ error: 'Failed to clear existing data' });
            }

            // Insert new data
            const stmt = db.prepare('INSERT INTO official_leaves (employee_id, date, reason) VALUES (?, ?, ?)');
            
            data.forEach(row => {
                const employeeId = row.employeeId || row.employee_id || row.id;
                const date = row.date || '';
                const reason = row.reason || '';
                
                stmt.run(employeeId, date, reason);
            });
            
            stmt.finalize();
            res.json({ success: true, count: data.length });
        });
    } catch (error) {
        console.error('Error processing official leaves file:', error);
        res.status(500).json({ error: 'Failed to process file' });
    }
});

// API endpoint to calculate presence
app.get('/api/calculate', (req, res) => {
    // Get all unique dates from card_logins to determine total working days
    db.all('SELECT DISTINCT date FROM card_logins', [], (err, dates) => {
        if (err) {
            return res.status(500).json({ error: 'Failed to retrieve dates' });
        }

        const totalWorkingDays = dates.length;

        // Query to get employee presence data
        const query = `
            SELECT 
                e.name, 
                e.directorate, 
                e.department,
                e.employee_id,
                (SELECT COUNT(DISTINCT date) FROM card_logins WHERE employee_id = e.employee_id) AS login_count,
                (SELECT COUNT(DISTINCT date) FROM official_leaves WHERE employee_id = e.employee_id) AS leave_count
            FROM 
                employees e
        `;

        db.all(query, [], (err, rows) => {
            if (err) {
                return res.status(500).json({ error: 'Failed to calculate presence' });
            }

            // Calculate days at home for each employee
            const results = rows.map(row => {
                const daysHome = Math.max(0, totalWorkingDays - row.login_count - row.leave_count);
                
                // Calculate building usage percentage
                let buildingUsage = 'N/A';
                if (totalWorkingDays > 0) {
                    const usageDays = row.login_count;
                    const availableDays = totalWorkingDays - row.leave_count;
                    
                    if (availableDays > 0) {
                        const percentage = (usageDays / availableDays) * 100;
                        buildingUsage = percentage.toFixed(1) + '%';
                    }
                }
                
                return {
                    name: row.name,
                    directorate: row.directorate,
                    department: row.department,
                    loginCount: row.login_count,
                    leaveCount: row.leave_count,
                    daysHome: daysHome,
                    buildingUsage: buildingUsage
                };
            });

            res.json({
                totalWorkingDays: totalWorkingDays,
                employees: results
            });
        });
    });
});

// Start the server
app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});