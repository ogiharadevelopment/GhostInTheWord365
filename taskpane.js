/* global Office */

// グローバル変数
let selectedFormat = null;
let savedFormats = {};

// Office.jsの初期化
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.addEventListener("DOMContentLoaded", initializeApp);
    }
});

// アプリケーションの初期化
function initializeApp() {
    // イベントリスナーの設定
    document.getElementById("save-format").addEventListener("click", saveFormat);
    document.getElementById("apply-format").addEventListener("click", applyFormat);
    document.getElementById("clear-format").addEventListener("click", clearFormat);
    
    // 保存された書式の読み込み
    loadSavedFormats();
    
    // 選択変更の監視
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    
    // 初期表示
    updateCurrentFormat();
}

// 選択変更時の処理
function onSelectionChanged() {
    updateCurrentFormat();
}

// 現在の書式を更新
function updateCurrentFormat() {
    Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            const font = selection.font;
            const paragraph = selection.paragraphs.getFirst();
            
            // 書式情報を読み込み
            font.load('name, size, bold, italic, color, underline, highlightColor');
            paragraph.load('alignment, leftIndent, rightIndent, lineSpacing, spaceAfter, spaceBefore');
            
            await context.sync();
            
            // 書式情報を取得
            const formatInfo = {
                font: {
                    name: font.name,
                    size: font.size,
                    bold: font.bold,
                    italic: font.italic,
                    color: font.color,
                    underline: font.underline,
                    highlightColor: font.highlightColor
                },
                paragraph: {
                    alignment: paragraph.alignment,
                    leftIndent: paragraph.leftIndent,
                    rightIndent: paragraph.rightIndent,
                    lineSpacing: paragraph.lineSpacing,
                    spaceAfter: paragraph.spaceAfter,
                    spaceBefore: paragraph.spaceBefore
                }
            };
            
            // 現在の書式を表示
            displayCurrentFormat(formatInfo);
            
        } catch (error) {
            console.error('書式取得エラー:', error);
            showMessage('書式の取得に失敗しました', 'error');
        }
    });
}

// 現在の書式を表示
function displayCurrentFormat(formatInfo) {
    const formatDisplay = document.getElementById('current-format');
    
    if (!formatInfo) {
        formatDisplay.innerHTML = '<p>テキストを選択して書式を確認してください</p>';
        return;
    }
    
    const font = formatInfo.font;
    const paragraph = formatInfo.paragraph;
    
    // 配置の日本語表示
    const alignmentText = getAlignmentText(paragraph.alignment);
    
    const formatText = `
        <div style="font-family: ${font.name}; font-size: ${font.size}px; 
                    font-weight: ${font.bold ? 'bold' : 'normal'}; 
                    font-style: ${font.italic ? 'italic' : 'normal'};
                    color: ${font.color}; text-align: ${alignmentText};">
            <strong>フォント:</strong> ${font.name}<br>
            <strong>サイズ:</strong> ${font.size}px<br>
            <strong>太字:</strong> ${font.bold ? 'ON' : 'OFF'}<br>
            <strong>斜体:</strong> ${font.italic ? 'ON' : 'OFF'}<br>
            <strong>色:</strong> ${font.color}<br>
            <strong>下線:</strong> ${font.underline}<br>
            <strong>配置:</strong> ${alignmentText}<br>
            <strong>左インデント:</strong> ${paragraph.leftIndent}pt<br>
            <strong>右インデント:</strong> ${paragraph.rightIndent}pt<br>
            <strong>行間:</strong> ${paragraph.lineSpacing}pt
        </div>
    `;
    
    formatDisplay.innerHTML = formatText;
}

// 配置の日本語表示を取得
function getAlignmentText(alignment) {
    switch (alignment) {
        case 'Left': return '左揃え';
        case 'Center': return '中央揃え';
        case 'Right': return '右揃え';
        case 'Justified': return '両端揃え';
        default: return alignment;
    }
}

// 書式を保存
function saveFormat() {
    const formatName = document.getElementById('format-name').value.trim();
    
    if (!formatName) {
        showMessage('書式名を入力してください', 'error');
        return;
    }
    
    Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            const font = selection.font;
            const paragraph = selection.paragraphs.getFirst();
            
            // 書式情報を読み込み
            font.load('name, size, bold, italic, color, underline, highlightColor');
            paragraph.load('alignment, leftIndent, rightIndent, lineSpacing, spaceAfter, spaceBefore');
            
            await context.sync();
            
            // 書式情報を保存
            const formatData = {
                font: {
                    name: font.name,
                    size: font.size,
                    bold: font.bold,
                    italic: font.italic,
                    color: font.color,
                    underline: font.underline,
                    highlightColor: font.highlightColor
                },
                paragraph: {
                    alignment: paragraph.alignment,
                    leftIndent: paragraph.leftIndent,
                    rightIndent: paragraph.rightIndent,
                    lineSpacing: paragraph.lineSpacing,
                    spaceAfter: paragraph.spaceAfter,
                    spaceBefore: paragraph.spaceBefore
                },
                timestamp: new Date().toISOString()
            };
            
            // ローカルストレージに保存
            savedFormats[formatName] = formatData;
            localStorage.setItem('savedFormats', JSON.stringify(savedFormats));
            
            // UIを更新
            updateSavedFormatsList();
            document.getElementById('format-name').value = '';
            
            showMessage(`書式「${formatName}」を保存しました`, 'success');
            
        } catch (error) {
            console.error('書式保存エラー:', error);
            showMessage('書式の保存に失敗しました', 'error');
        }
    });
}

// 書式を適用
function applyFormat() {
    if (!selectedFormat) {
        showMessage('適用する書式を選択してください', 'error');
        return;
    }
    
    Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            const font = selection.font;
            const paragraph = selection.paragraphs.getFirst();
            
            // 書式を適用
            font.name = selectedFormat.font.name;
            font.size = selectedFormat.font.size;
            font.bold = selectedFormat.font.bold;
            font.italic = selectedFormat.font.italic;
            font.color = selectedFormat.font.color;
            font.underline = selectedFormat.font.underline;
            font.highlightColor = selectedFormat.font.highlightColor;
            
            paragraph.alignment = selectedFormat.paragraph.alignment;
            paragraph.leftIndent = selectedFormat.paragraph.leftIndent;
            paragraph.rightIndent = selectedFormat.paragraph.rightIndent;
            paragraph.lineSpacing = selectedFormat.paragraph.lineSpacing;
            paragraph.spaceAfter = selectedFormat.paragraph.spaceAfter;
            paragraph.spaceBefore = selectedFormat.paragraph.spaceBefore;
            
            await context.sync();
            
            showMessage('書式を適用しました', 'success');
            
        } catch (error) {
            console.error('書式適用エラー:', error);
            showMessage('書式の適用に失敗しました', 'error');
        }
    });
}

// 書式をクリア
function clearFormat() {
    Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            const font = selection.font;
            const paragraph = selection.paragraphs.getFirst();
            
            // 書式をクリア
            font.name = 'Calibri';
            font.size = 11;
            font.bold = false;
            font.italic = false;
            font.color = 'black';
            font.underline = 'None';
            font.highlightColor = 'NoColor';
            
            paragraph.alignment = 'Left';
            paragraph.leftIndent = 0;
            paragraph.rightIndent = 0;
            paragraph.lineSpacing = 0;
            paragraph.spaceAfter = 0;
            paragraph.spaceBefore = 0;
            
            await context.sync();
            
            showMessage('書式をクリアしました', 'success');
            
        } catch (error) {
            console.error('書式クリアエラー:', error);
            showMessage('書式のクリアに失敗しました', 'error');
        }
    });
}

// 保存された書式を読み込み
function loadSavedFormats() {
    try {
        const saved = localStorage.getItem('savedFormats');
        if (saved) {
            savedFormats = JSON.parse(saved);
            updateSavedFormatsList();
        }
    } catch (error) {
        console.error('保存された書式の読み込みエラー:', error);
    }
}

// 保存された書式一覧を更新
function updateSavedFormatsList() {
    const savedFormatsList = document.getElementById('saved-formats');
    
    if (Object.keys(savedFormats).length === 0) {
        savedFormatsList.innerHTML = '<p>保存された書式がここに表示されます</p>';
        return;
    }
    
    let html = '';
    for (const [name, format] of Object.entries(savedFormats)) {
        const date = new Date(format.timestamp).toLocaleDateString('ja-JP');
        html += `
            <div class="format-item" data-format-name="${name}">
                <div class="format-item-name">${name}</div>
                <div class="format-item-preview">${format.font.name} ${format.font.size}px - ${getAlignmentText(format.paragraph.alignment)} (${date})</div>
            </div>
        `;
    }
    
    savedFormatsList.innerHTML = html;
    
    // クリックイベントを追加
    const formatItems = savedFormatsList.querySelectorAll('.format-item');
    formatItems.forEach(item => {
        item.addEventListener('click', () => {
            // 選択状態を更新
            formatItems.forEach(i => i.classList.remove('selected'));
            item.classList.add('selected');
            
            // 選択された書式を設定
            const formatName = item.dataset.formatName;
            selectedFormat = savedFormats[formatName];
            
            // 適用ボタンを有効化
            document.getElementById('apply-format').disabled = false;
            
            // 書式詳細を表示
            displayFormatDetails(selectedFormat);
        });
    });
}

// 書式詳細を表示
function displayFormatDetails(format) {
    const formatDetails = document.getElementById('format-details');
    
    if (!format) {
        formatDetails.innerHTML = '<p>書式の詳細情報がここに表示されます</p>';
        return;
    }
    
    const details = JSON.stringify(format, null, 2);
    formatDetails.innerHTML = `<pre>${details}</pre>`;
}

// メッセージを表示
function showMessage(message, type) {
    // 既存のメッセージを削除
    const existingMessage = document.querySelector('.error-message, .success-message');
    if (existingMessage) {
        existingMessage.remove();
    }
    
    // 新しいメッセージを作成
    const messageDiv = document.createElement('div');
    messageDiv.className = type === 'error' ? 'error-message' : 'success-message';
    messageDiv.textContent = message;
    
    // メッセージを表示
    const saveButton = document.getElementById('save-format');
    saveButton.parentNode.appendChild(messageDiv);
    
    // 3秒後にメッセージを削除
    setTimeout(() => {
        if (messageDiv.parentNode) {
            messageDiv.remove();
        }
    }, 3000);
}
