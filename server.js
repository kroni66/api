const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');
const multer = require('multer');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Create uploads directory if it doesn't exist
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}

// Multer setup for .xlsx file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadsDir);
    },
    filename: (req, file, cb) => {
        const uniqueName = `uploaded_${uuidv4()}.xlsx`;
        cb(null, uniqueName);
    }
});
const upload = multer({
    storage,
    fileFilter: (req, file, cb) => {
        if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            cb(null, true);
        } else {
            cb(new Error('Only .xlsx files are allowed!'));
        }
    }
});

// Store for generated files (in production, use a database)
const fileStore = new Map();

/**
 * Process ChatGPT payload and generate Excel file
 * @param {Object} payload - ChatGPT payload data
 * @returns {String} - File path of generated Excel file
 */
async function processPayloadToExcel(payload) {
    const workbook = new ExcelJS.Workbook();
    
    // Add metadata
    workbook.creator = 'ChatGPT Excel API';
    workbook.created = new Date();
    
    // Main worksheet for ChatGPT content
    const worksheet = workbook.addWorksheet('ChatGPT Data');
    
    // Configure columns based on payload structure
    if (payload.conversations && Array.isArray(payload.conversations)) {
        // Handle conversation format
        worksheet.columns = [
            { header: 'Timestamp', key: 'timestamp', width: 20 },
            { header: 'Role', key: 'role', width: 15 },
            { header: 'Content', key: 'content', width: 80 },
            { header: 'Token Count', key: 'tokens', width: 15 }
        ];
        
        // Add conversation data
        payload.conversations.forEach((conv, index) => {
            worksheet.addRow({
                timestamp: conv.timestamp || new Date().toISOString(),
                role: conv.role || 'unknown',
                content: conv.content || conv.message || '',
                tokens: conv.token_count || conv.tokens || 0
            });
        });
    } else if (payload.data && Array.isArray(payload.data)) {
        // Handle generic data array format
        if (payload.data.length > 0) {
            const firstItem = payload.data[0];
            const headers = Object.keys(firstItem);
            
            worksheet.columns = headers.map(header => ({
                header: header.charAt(0).toUpperCase() + header.slice(1),
                key: header,
                width: 20
            }));
            
            payload.data.forEach(item => {
                worksheet.addRow(item);
            });
        }
    } else {
        // Handle single object or unknown format
        worksheet.columns = [
            { header: 'Property', key: 'property', width: 30 },
            { header: 'Value', key: 'value', width: 50 }
        ];
        
        // Flatten the payload object
        const flattenObject = (obj, prefix = '') => {
            const rows = [];
            for (const [key, value] of Object.entries(obj)) {
                const fullKey = prefix ? `${prefix}.${key}` : key;
                if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
                    rows.push(...flattenObject(value, fullKey));
                } else {
                    rows.push({
                        property: fullKey,
                        value: Array.isArray(value) ? JSON.stringify(value) : String(value)
                    });
                }
            }
            return rows;
        };
        
        const flatData = flattenObject(payload);
        flatData.forEach(row => worksheet.addRow(row));
    }
    
    // Style the header row
    worksheet.getRow(1).eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '366092' }
        };
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    
    // Add borders to all cells
    worksheet.eachRow((row) => {
        row.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    });
    
    // Auto-fit columns
    worksheet.columns.forEach(column => {
        if (column.width < 10) column.width = 10;
        if (column.width > 100) column.width = 100;
    });
    
    // Generate unique filename
    const filename = `chatgpt_export_${uuidv4()}.xlsx`;
    const filepath = path.join(uploadsDir, filename);
    
    // Save the workbook
    await workbook.xlsx.writeFile(filepath);
    
    return { filename, filepath };
}

/**
 * API endpoint to upload an .xlsx file
 */
app.post('/api/upload-xlsx', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({
                error: 'No file uploaded',
                message: 'Please upload a .xlsx file'
            });
        }

        const fileId = uuidv4();
        const fileInfo = {
            filename: req.file.filename,
            filepath: req.file.path,
            createdAt: new Date().toISOString()
        };
        fileStore.set(fileId, fileInfo);

        res.status(201).json({
            success: true,
            message: 'File uploaded successfully',
            fileId,
            filename: req.file.filename,
            downloadUrl: `/api/download/${fileId}`
        });
    } catch (error) {
        console.error('Error uploading file:', error);
        res.status(500).json({
            error: 'Internal server error',
            message: 'Failed to upload file'
        });
    }
});

/**
 * API endpoint to download generated Excel file
 */
app.get('/api/download/:fileId', (req, res) => {
    try {
        const { fileId } = req.params;
        const { attachment } = req.query;
        
        // Get file info from store
        const fileInfo = fileStore.get(fileId);
        
        if (!fileInfo) {
            return res.status(404).json({
                error: 'File not found',
                message: 'The requested file does not exist or has expired'
            });
        }
        
        // Check if file exists on disk
        if (!fs.existsSync(fileInfo.filepath)) {
            // Remove from store if file doesn't exist
            fileStore.delete(fileId);
            return res.status(404).json({
                error: 'File not found',
                message: 'The file has been removed from the server'
            });
        }
        
        // Set appropriate headers
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        
        if (attachment === 'true') {
            res.setHeader('Content-Disposition', `attachment; filename="${fileInfo.filename}"`);
        } else {
            res.setHeader('Content-Disposition', `inline; filename="${fileInfo.filename}"`);
        }
        
        // Stream the file
        const fileStream = fs.createReadStream(fileInfo.filepath);
        fileStream.pipe(res);
        
        fileStream.on('error', (error) => {
            console.error('Error streaming file:', error);
            if (!res.headersSent) {
                res.status(500).json({
                    error: 'File streaming error',
                    message: 'Failed to download the file'
                });
            }
        });
        
    } catch (error) {
        console.error('Error downloading file:', error);
        res.status(500).json({
            error: 'Internal server error',
            message: 'Failed to download file',
            details: error.message
        });
    }
});

/**
 * API endpoint to list all generated files
 */
app.get('/api/files', (req, res) => {
    try {
        const files = Array.from(fileStore.entries()).map(([fileId, fileInfo]) => ({
            fileId,
            filename: fileInfo.filename,
            createdAt: fileInfo.createdAt,
            downloadUrl: `/api/download/${fileId}`,
            size: fs.existsSync(fileInfo.filepath) ? fs.statSync(fileInfo.filepath).size : 0
        }));
        
        res.json({
            success: true,
            files,
            count: files.length
        });
    } catch (error) {
        console.error('Error listing files:', error);
        res.status(500).json({
            error: 'Internal server error',
            message: 'Failed to list files'
        });
    }
});

/**
 * API endpoint to delete a generated file
 */
app.delete('/api/files/:fileId', (req, res) => {
    try {
        const { fileId } = req.params;
        const fileInfo = fileStore.get(fileId);
        
        if (!fileInfo) {
            return res.status(404).json({
                error: 'File not found',
                message: 'The requested file does not exist'
            });
        }
        
        // Delete file from disk if it exists
        if (fs.existsSync(fileInfo.filepath)) {
            fs.unlinkSync(fileInfo.filepath);
        }
        
        // Remove from store
        fileStore.delete(fileId);
        
        res.json({
            success: true,
            message: 'File deleted successfully'
        });
        
    } catch (error) {
        console.error('Error deleting file:', error);
        res.status(500).json({
            error: 'Internal server error',
            message: 'Failed to delete file'
        });
    }
});

/**
 * Health check endpoint
 */
app.get('/api/health', (req, res) => {
    res.json({
        status: 'OK',
        timestamp: new Date().toISOString(),
        uptime: process.uptime(),
        version: '1.0.0'
    });
});

/**
 * Root endpoint with API documentation
 */
app.get('/', (req, res) => {
    res.json({
        message: 'ChatGPT Excel API',
        version: '1.0.0',
        endpoints: {
            'POST /api/process-chatgpt': 'Process ChatGPT payload and generate Excel file',
            'POST /api/upload-xlsx': 'Upload a .xlsx file',
            'GET /api/download/:fileId': 'Download generated Excel file',
            'GET /api/files': 'List all generated files',
            'DELETE /api/files/:fileId': 'Delete a generated file',
            'GET /api/health': 'Health check'
        },
        documentation: {
            'ChatGPT payload format examples': {
                'Conversation format': {
                    conversations: [
                        {
                            timestamp: '2024-01-01T12:00:00Z',
                            role: 'user',
                            content: 'Hello, how are you?',
                            tokens: 5
                        },
                        {
                            timestamp: '2024-01-01T12:00:05Z',
                            role: 'assistant',
                            content: 'I am doing well, thank you!',
                            tokens: 8
                        }
                    ]
                },
                'Data array format': {
                    data: [
                        { name: 'John', age: 30, city: 'New York' },
                        { name: 'Jane', age: 25, city: 'Los Angeles' }
                    ]
                },
                'Generic object format': {
                    title: 'My ChatGPT Session',
                    user: 'john_doe',
                    session_id: 'abc123',
                    messages: ['Hello', 'How are you?'],
                    metadata: { tokens_used: 50, duration: '5 minutes' }
                }
            }
        }
    });
});

// Error handling middleware
app.use((error, req, res, next) => {
    console.error('Unhandled error:', error);
    res.status(500).json({
        error: 'Internal server error',
        message: 'An unexpected error occurred'
    });
});

// Start server
app.listen(PORT, () => {
    console.log(`🚀 ChatGPT Excel API server is running on port ${PORT}`);
    console.log(`📊 API Documentation: http://localhost:${PORT}/`);
    console.log(`💚 Health Check: http://localhost:${PORT}/api/health`);
    console.log(`📁 Upload directory: ${uploadsDir}`);
});

module.exports = app; 