const express = require('express');
const path = require('path');
const app = express();
const port = 3000;

// 静的ファイルの提供
app.use(express.static('.'));

// CORSヘッダーの設定
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
    
    if (req.method === 'OPTIONS') {
        res.sendStatus(200);
    } else {
        next();
    }
});

// ルートパス
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'taskpane.html'));
});

// サーバー起動
app.listen(port, () => {
    console.log(`Word 365 書式管理アドイン サーバーが起動しました: http://localhost:${port}`);
    console.log('Word 365でアドインをサイドロードしてください。');
});
