/* global Office */

// グローバル変数
let currentFormat = null;
let savedFormats = {};
let currentLanguage = 'ja';
let currentFontSize = 12;
let currentLineSpacing = 1.0;
let isWideMode = true;
let selectedArea = null;
let savedCursorPosition = null; // カーソル位置を保存
let continuousMode = false; // 連続モード
let continuousFormat = null; // 連続適用用の書式
let isMouseOverSaveArea = false; // SAVEエリアのマウスオーバー状態

// 多言語対応テキスト
const texts = {
    ja: {
        appTitle: '書式管理',
        currentFormatTitle: '現在の書式',
        noSelectionText: 'テキストを選択してください',
        saveLabel: 'SAVE',
        saveInstruction: 'キーを押して保存',
        loadLabel: 'LOAD',
        loadInstruction: 'キーを押して適用',
        savedFormatsTitle: '保存された書式',
        noSavedFormatsText: '保存された書式はありません',
        keyGuideTitle: 'キーガイド',
        keyGuideText: '保存された書式にマウスオーバーしてキーを押すと書式を適用します',
        fontLabel: 'フォント',
        continuousLabel: '連続',
        formatSaved: '書式を保存しました',
        formatApplied: '書式を適用しました',
        formatNotFound: '保存された書式が見つかりません',
        noTextSelected: 'テキストが選択されていません',
        widthToggle: '幅: 300px',
        widthToggleNarrow: '幅: 100px',
        deleteConfirm: (key) => `書式 "${key}" を削除しますか？`,
        savedFormatsInstruction: 'マウスオーバーしてキーを押すと適用',
        continuousModeOn: 'ON',
        continuousModeOff: 'OFF',
        continuousModeEnabled: '連続モード有効',
        continuousModeDisabled: '連続モード無効',
        continuousFormatSaved: '連続適用用書式を保存しました',
        japanese: '日本語',
        english: 'English'
    },
    en: {
        appTitle: 'Format Manager',
        currentFormatTitle: 'Current Format',
        noSelectionText: 'Please select text',
        saveLabel: 'SAVE',
        saveInstruction: 'Press key to save',
        loadLabel: 'LOAD',
        loadInstruction: 'Press key to apply',
        savedFormatsTitle: 'Saved Formats',
        noSavedFormatsText: 'No saved formats',
        keyGuideTitle: 'Key Guide',
        keyGuideText: 'Mouse over a saved format and press a key to apply it',
        fontLabel: 'Font',
        continuousLabel: 'Continuous',
        formatSaved: 'Format saved',
        formatApplied: 'Format applied',
        formatNotFound: 'Saved format not found',
        noTextSelected: 'No text selected',
        widthToggle: 'Width: 300px',
        widthToggleNarrow: 'Width: 100px',
        deleteConfirm: (key) => `Delete format "${key}"?`,
        savedFormatsInstruction: 'Mouse over and press key to apply',
        continuousModeOn: 'ON',
        continuousModeOff: 'OFF',
        continuousModeEnabled: 'Continuous mode enabled',
        continuousModeDisabled: 'Continuous mode disabled',
        continuousFormatSaved: 'Continuous format saved',
        japanese: '日本語',
        english: 'English'
    }
};

// Office.jsの初期化
Office.onReady((info) => {
    console.log('=== Office.onReady called ===');
    console.log('Info object:', JSON.stringify(info, null, 2));
    console.log('Host type:', info.host);
    console.log('Platform:', info.platform);
    
    if (info.host === Office.HostType.Word) {
        console.log('✅ Word host detected - proceeding with initialization');
        console.log('DOM ready state:', document.readyState);
        
        // DOMContentLoadedイベントに依存せず、直接初期化を試行
        if (document.readyState === 'loading') {
            console.log('DOM still loading, waiting for DOMContentLoaded');
            document.addEventListener("DOMContentLoaded", initializeApp);
        } else {
            console.log('DOM already ready, initializing immediately');
            // 少し遅延してから初期化（DOM要素が確実に存在するように）
            setTimeout(initializeApp, 100);
        }
        
        // フォールバック: 3秒後に強制初期化
        setTimeout(() => {
            console.log('Fallback initialization after 3 seconds');
            initializeApp();
        }, 3000);
    } else {
        console.log('❌ Non-Word host detected:', info.host);
        console.log('Expected:', Office.HostType.Word);
    }
}).catch(error => {
    console.error('❌ Office.onReady error:', error);
});

// アプリケーションの初期化
function initializeApp() {
    console.log('=== initializeApp called ===');
    console.log('DOM ready state:', document.readyState);
    console.log('Current time:', new Date().toISOString());
    console.log('Document body exists:', !!document.body);
    console.log('Document head exists:', !!document.head);
    
    // 重複初期化を防ぐ
    if (window.appInitialized) {
        console.log('App already initialized, skipping');
        return;
    }
    window.appInitialized = true;
    
    try {
        console.log('Step 1: Word API availability check');
        // Word APIの可用性チェック
        checkWordAPIAvailability();
        
        console.log('Step 2: Language loading');
        // 言語設定の読み込み
        loadLanguage();
        
        console.log('Step 3: UI update');
        // UIの初期化
        updateUI();
        
        console.log('Step 4: Event listeners setup');
        // イベントリスナーの設定
        setupEventListeners();
        
        // 要素の存在確認
        console.log('=== Element existence check ===');
        const saveArea = document.getElementById('save-area');
        const fontControl = document.getElementById('font-control');
        const continuousControl = document.getElementById('continuous-control');
        const langJa = document.getElementById('lang-ja');
        const langEn = document.getElementById('lang-en');
        
        console.log('Save area found:', !!saveArea);
        console.log('Font control found:', !!fontControl);
        console.log('Continuous control found:', !!continuousControl);
        console.log('Japanese button found:', !!langJa);
        console.log('English button found:', !!langEn);
        
        if (!saveArea || !fontControl || !continuousControl) {
            console.error('❌ Critical elements missing - retrying in 500ms');
            window.appInitialized = false; // リトライのためにフラグをリセット
            setTimeout(initializeApp, 500);
            return;
        }
        
        console.log('Step 5: Saved formats loading');
        // 保存された書式の読み込み
        loadSavedFormats();
        
        console.log('Step 6: Selection change handler');
        // 選択変更の監視
        try {
            Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
            console.log('✅ Selection change handler added');
        } catch (error) {
            console.error('❌ Failed to add selection change handler:', error);
        }
        
        console.log('Step 7: Initial format update');
        // 初期表示
        updateCurrentFormat();
        
        console.log('Step 8: Synthetic click setup');
        // 疑似クリックイベントの設定
        setupSyntheticClick();
        
        console.log('Step 9: Final UI update');
        // 最終的なUI更新
        updateSavedFormatsList();
        
        console.log('Step 10: Initialize display values');
        // 初期表示値を設定
        updateFontSizeDisplay();
        updateContinuousDisplay();
        
        console.log('✅ App initialization completed successfully');
        console.log('=== Initialization Summary ===');
        console.log('All steps completed without errors');
        console.log('Ready for user interaction');
    } catch (error) {
        console.error('❌ App initialization error:', error);
        console.error('Error stack:', error.stack);
        window.appInitialized = false; // エラー時はフラグをリセット
    }
}

// イベントリスナーの設定
function setupEventListeners() {
    console.log('=== setupEventListeners called ===');
    
    try {
        // 言語切り替え
        const langJa = document.getElementById('lang-ja');
        const langEn = document.getElementById('lang-en');
        
        if (langJa) {
            langJa.addEventListener('click', () => setLanguage('ja'));
            console.log('✅ Japanese language button event added');
        } else {
            console.error('❌ Japanese language button not found');
        }
        
        if (langEn) {
            langEn.addEventListener('click', () => setLanguage('en'));
            console.log('✅ English language button event added');
        } else {
            console.error('❌ English language button not found');
        }
    
        // コントロール領域のイベント
        const saveArea = document.getElementById('save-area');
        const fontControl = document.getElementById('font-control');
        const continuousControl = document.getElementById('continuous-control');
        const widthToggle = document.getElementById('width-toggle');
        
        if (saveArea) {
            console.log('✅ Save area found');
            // マウスイベント
            saveArea.addEventListener('mouseenter', async (e) => {
                console.log('🖱️ Save area mouseenter');
                e.preventDefault();
                isMouseOverSaveArea = true; // マウスオーバー状態を設定
                await saveCursorPosition(); // カーソル位置を保存
                selectArea('save');
                // フォーカスを確実に取得
                setTimeout(() => {
                    saveArea.focus();
                }, 10);
            });
            
            saveArea.addEventListener('mouseleave', async (e) => {
                console.log('🖱️ Save area mouseleave');
                isMouseOverSaveArea = false; // マウスオーバー状態を解除
                await restoreCursorPosition(); // カーソル位置を復元
            });
            
            console.log('✅ Save area mouseenter event added');
        } else {
            console.error('❌ Save area not found');
        }
        
        if (fontControl) {
            console.log('✅ Font control found');
            fontControl.addEventListener('mouseenter', async (e) => {
                console.log('🖱️ Font control mouseenter');
                e.preventDefault();
                await saveCursorPosition(); // カーソル位置を保存
                selectArea('font');
                setTimeout(() => {
                    fontControl.focus();
                    fontControl.click();
                }, 10);
            });
            
            fontControl.addEventListener('mouseleave', async (e) => {
                console.log('🖱️ Font control mouseleave');
                await restoreCursorPosition(); // カーソル位置を復元
            });
            
            fontControl.addEventListener('wheel', handleFontWheel);
            console.log('✅ Font control events added');
        } else {
            console.error('❌ Font control not found');
        }
        
        if (continuousControl) {
            console.log('✅ Continuous control found');
            continuousControl.addEventListener('mouseenter', async (e) => {
                console.log('🖱️ Continuous control mouseenter');
                e.preventDefault();
                await saveCursorPosition(); // カーソル位置を保存
                selectArea('continuous');
                setTimeout(() => {
                    continuousControl.focus();
                }, 10);
            });
            
            continuousControl.addEventListener('mouseleave', async (e) => {
                console.log('🖱️ Continuous control mouseleave');
                await restoreCursorPosition(); // カーソル位置を復元
            });
            
            continuousControl.addEventListener('keydown', handleKeyPress);
            console.log('✅ Continuous control events added');
        } else {
            console.error('❌ Continuous control not found');
        }
        
        if (widthToggle) {
            console.log('✅ Width toggle found');
            widthToggle.addEventListener('click', toggleWidth);
            console.log('✅ Width toggle event added');
        } else {
            console.error('❌ Width toggle not found');
        }
    
        // フォーカスイベント
        if (saveArea) {
            saveArea.addEventListener('focus', () => selectArea('save'));
            console.log('✅ Save area focus event added');
        }
        if (fontControl) {
            fontControl.addEventListener('focus', () => selectArea('font'));
            console.log('✅ Font control focus event added');
        }
        if (continuousControl) {
            continuousControl.addEventListener('focus', () => selectArea('continuous'));
            console.log('✅ Continuous control focus event added');
        }
        
        // キーボードイベント
        if (saveArea) {
            saveArea.addEventListener('keydown', handleKeyPress);
            console.log('✅ Save area keydown event added');
        }
        if (fontControl) {
            fontControl.addEventListener('keydown', handleKeyPress);
            console.log('✅ Font control keydown event added');
        }
        if (continuousControl) {
            continuousControl.addEventListener('keydown', handleKeyPress);
            console.log('✅ Continuous control keydown event added');
        }
        
        // クリックイベント（フォーカス用）
        if (saveArea) {
            saveArea.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                saveArea.focus();
            });
            console.log('✅ Save area click event added');
        }
        if (fontControl) {
            fontControl.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                fontControl.focus();
            });
            console.log('✅ Font control click event added');
        }
        if (continuousControl) {
            continuousControl.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                // クリックのみでオン/オフ切り替え
                toggleContinuousMode();
                continuousControl.focus();
            });
            console.log('✅ Continuous control click event added');
        }
        
        // マウスリーブイベント（フォーカスを維持）
        if (saveArea) {
            saveArea.addEventListener('mouseleave', () => {
                // フォーカスを維持
            });
            console.log('✅ Save area mouseleave event added');
        }
        if (fontControl) {
            fontControl.addEventListener('mouseleave', () => {
                // フォーカスを維持
            });
            console.log('✅ Font control mouseleave event added');
        }
        if (continuousControl) {
            continuousControl.addEventListener('mouseleave', () => {
                // フォーカスを維持
            });
            console.log('✅ Continuous control mouseleave event added');
        }
        
        console.log('✅ setupEventListeners completed successfully');
    } catch (error) {
        console.error('❌ setupEventListeners error:', error);
        console.error('Error stack:', error.stack);
    }
}

// 言語設定
function setLanguage(lang) {
    currentLanguage = lang;
    localStorage.setItem('formatManagerLanguage', lang);
    updateUI();
    
    // 言語ボタンの状態更新
    document.querySelectorAll('.lang-btn').forEach(btn => btn.classList.remove('active'));
    document.getElementById(`lang-${lang}`).classList.add('active');
}

// 言語設定の読み込み
function loadLanguage() {
    const savedLang = localStorage.getItem('formatManagerLanguage');
    if (savedLang && texts[savedLang]) {
        currentLanguage = savedLang;
    }
}

// UIの更新
function updateUI() {
    const t = texts[currentLanguage];
    
    // 要素の存在確認をしてから更新
    const elements = {
        'app-title': t.appTitle,
        'current-format-title': t.currentFormatTitle,
        'no-selection-text': t.noSelectionText,
        'save-label': t.saveLabel,
        'save-instruction': t.saveInstruction,
        'saved-formats-title': t.savedFormatsTitle,
        'no-saved-formats-text': t.noSavedFormatsText,
        'key-guide-title': t.keyGuideTitle,
        'key-guide-text': t.keyGuideText,
        'font-label': t.fontLabel,
        'continuous-label': t.continuousLabel,
        'width-toggle': t.widthToggle,
        'saved-formats-instruction': t.savedFormatsInstruction,
        'lang-ja': t.japanese,
        'lang-en': t.english
    };
    
    for (const [id, text] of Object.entries(elements)) {
        const element = document.getElementById(id);
        if (element) {
            element.textContent = text;
        } else {
            console.warn(`Element with id '${id}' not found`);
        }
    }
}

        // カーソル位置を保存（簡素化版）
        async function saveCursorPosition() {
            try {
                if (typeof Word === 'undefined') {
                    console.log('❌ Word API not available for cursor position save');
                    return;
                }
                
                console.log('💾 Starting cursor position save...');
                
                await Word.run(async (context) => {
                    const selection = context.document.getSelection();
                    
                    // 基本的な情報を取得
                    selection.load('text, isEmpty');
                    await context.sync();
                    
                    console.log('📝 Selection info:', {
                        text: selection.text,
                        isEmpty: selection.isEmpty,
                        textLength: selection.text ? selection.text.length : 0
                    });
                    
                    // Word Onlineでは位置情報の取得が制限されているため、
                    // 選択されたテキストのみを保存
                    if (selection.text && selection.text.trim() !== '') {
                        savedCursorPosition = {
                            type: 'selection',
                            text: selection.text,
                            timestamp: new Date().toISOString()
                        };
                        console.log('✅ Selection text saved:', savedCursorPosition);
                    } else {
                        console.log('ℹ️ No text selected - cursor position save skipped');
                        savedCursorPosition = null;
                    }
                });
            } catch (error) {
                console.error('❌ Failed to save cursor position:', error);
                savedCursorPosition = null;
            }
        }

        // カーソル位置を復元（簡素化版）
        async function restoreCursorPosition() {
            try {
                if (!savedCursorPosition) {
                    console.log('ℹ️ No saved cursor position to restore');
                    return;
                }
                
                if (typeof Word === 'undefined') {
                    console.log('❌ Word API not available for cursor position restore');
                    return;
                }
                
                console.log('🔄 Starting cursor position restore...', savedCursorPosition);
                
                // Word Onlineでは位置情報の復元が制限されているため、
                // 選択されたテキストの検索のみを試行
                if (savedCursorPosition.type === 'selection' && savedCursorPosition.text) {
                    await Word.run(async (context) => {
                        const body = context.document.body;
                        body.load('text');
                        await context.sync();
                        
                        const documentText = body.text || '';
                        const searchText = savedCursorPosition.text;
                        
                        if (documentText.includes(searchText)) {
                            const startIndex = documentText.indexOf(searchText);
                            const endIndex = startIndex + searchText.length;
                            
                            const selection = context.document.getSelection();
                            selection.select(startIndex, endIndex);
                            await context.sync();
                            
                            console.log('✅ Selection restored by text search');
                        } else {
                            console.log('ℹ️ Saved text not found in document');
                        }
                    });
                } else {
                    console.log('ℹ️ No valid selection to restore');
                }
            } catch (error) {
                console.error('❌ Failed to restore cursor position:', error);
            }
        }

// 領域の選択
function selectArea(area) {
    selectedArea = area;
    
    // 視覚的フィードバック
    document.querySelectorAll('.action-area, .control-area').forEach(el => el.classList.remove('selected'));
    
    // 対応する要素にクラスを追加
    if (area === 'save') {
        const saveArea = document.getElementById('save-area');
        if (saveArea) saveArea.classList.add('selected');
    } else if (area === 'font') {
        const fontControl = document.getElementById('font-control');
        if (fontControl) fontControl.classList.add('selected');
    } else if (area === 'lineSpacing') {
        const lineSpacingControl = document.getElementById('line-spacing-control');
        if (lineSpacingControl) lineSpacingControl.classList.add('selected');
    }
}

// キー押下の処理
function handleKeyPress(event) {
    // 特殊キーは無視
    if (event.key === 'Tab' || event.key === 'Shift' || event.key === 'Control' ||
        event.key === 'Alt' || event.key === 'Meta' || event.key === 'CapsLock' ||
        event.key === 'Enter' || event.key === 'Escape' || event.key === 'ArrowUp' ||
        event.key === 'ArrowDown' || event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
        return;
    }
    
    event.preventDefault();
    event.stopPropagation();
    
    const key = event.key.toLowerCase();
    const targetId = event.currentTarget.id;
    
    console.log(`Key pressed: ${key} in ${targetId}`);
    
    if (targetId === 'save-area') {
        // SAVEエリアはマウスオーバー中のみキー入力を受け付ける
        if (isMouseOverSaveArea) {
            saveFormat(key);
        }
    } else if (targetId === 'font-control') {
        adjustFontSize(key);
    } else if (targetId === 'continuous-control') {
        // 連続ボタンは既存の保存された書式を連続適用用に設定
        setContinuousFormat(key);
    }
    
    // 視覚的フィードバック
    if (event.currentTarget && event.currentTarget.classList) {
        event.currentTarget.classList.add('pulse');
        setTimeout(() => {
            if (event.currentTarget && event.currentTarget.classList) {
                event.currentTarget.classList.remove('pulse');
            }
        }, 300);
    }
}

// 書式の保存
function saveFormat(key) {
    if (!currentFormat) {
        showMessage(texts[currentLanguage].noTextSelected, 'error');
        return;
    }
    
    try {
        savedFormats[key] = {
            ...currentFormat,
            timestamp: new Date().toISOString()
        };
        
        localStorage.setItem('savedFormats', JSON.stringify(savedFormats));
        updateSavedFormatsList();
        
        // 視覚的フィードバック
        const saveArea = document.getElementById('save-area');
        saveArea.classList.add('saved');
        setTimeout(() => saveArea.classList.remove('saved'), 1000);
        
        showMessage(`${key}: ${texts[currentLanguage].formatSaved}`, 'success');
        
    } catch (error) {
        console.error('書式保存エラー:', error);
        showMessage('書式の保存に失敗しました', 'error');
    }
}

        // 書式の適用（保存された書式から）
        function loadFormat(key) {
            if (!savedFormats[key]) {
                showMessage(texts[currentLanguage].formatNotFound, 'error');
                return;
            }

            Word.run(async (context) => {
                try {
                    const selection = context.document.getSelection();
                    const format = savedFormats[key];

                    // 選択範囲を確認
                    selection.load('text');
                    await context.sync();

                    console.log('🎨 Applying format:', {
                        key,
                        selectedText: selection.text,
                        hasSelection: selection.text && selection.text.trim() !== ''
                    });

                    // 書式を適用（選択されていない状態でも適用可能）
                    const font = selection.font;
                    const paragraph = selection.paragraphs.getFirst();

                    console.log('🎨 Applying format to selection:', {
                        hasSelection: selection.text && selection.text.trim() !== '',
                        selectedText: selection.text
                    });

                    // フォント書式を適用
                    if (format.font.name) {
                        font.name = format.font.name;
                        console.log('✅ Font name applied:', format.font.name);
                    }
                    if (format.font.size) {
                        font.size = format.font.size;
                        console.log('✅ Font size applied:', format.font.size);
                    }
                    if (format.font.bold !== undefined) {
                        font.bold = format.font.bold;
                        console.log('✅ Bold applied:', format.font.bold);
                    }
                    if (format.font.italic !== undefined) {
                        font.italic = format.font.italic;
                        console.log('✅ Italic applied:', format.font.italic);
                    }
                    if (format.font.color) {
                        font.color = format.font.color;
                        console.log('✅ Font color applied:', format.font.color);
                    }
                    if (format.font.underline !== undefined) {
                        font.underline = format.font.underline;
                        console.log('✅ Underline applied:', format.font.underline);
                    }
                    if (format.font.highlightColor) {
                        font.highlightColor = format.font.highlightColor;
                        console.log('✅ Highlight color applied:', format.font.highlightColor);
                    }

                    // 段落書式を適用
                    if (format.paragraph.alignment) {
                        paragraph.alignment = format.paragraph.alignment;
                        console.log('✅ Alignment applied:', format.paragraph.alignment);
                    }
                    if (format.paragraph.leftIndent !== undefined) {
                        paragraph.leftIndent = format.paragraph.leftIndent;
                        console.log('✅ Left indent applied:', format.paragraph.leftIndent);
                    }
                    if (format.paragraph.rightIndent !== undefined) {
                        paragraph.rightIndent = format.paragraph.rightIndent;
                        console.log('✅ Right indent applied:', format.paragraph.rightIndent);
                    }
                    if (format.paragraph.lineSpacing !== undefined) {
                        paragraph.lineSpacing = format.paragraph.lineSpacing;
                        console.log('✅ Line spacing applied:', format.paragraph.lineSpacing);
                    }
                    if (format.paragraph.spaceAfter !== undefined) {
                        paragraph.spaceAfter = format.paragraph.spaceAfter;
                        console.log('✅ Space after applied:', format.paragraph.spaceAfter);
                    }
                    if (format.paragraph.spaceBefore !== undefined) {
                        paragraph.spaceBefore = format.paragraph.spaceBefore;
                        console.log('✅ Space before applied:', format.paragraph.spaceBefore);
                    }

                    await context.sync();

                    // アドイン内の書式表示を更新
                    await updateCurrentFormatDisplay(format);

                    const message = selection.text && selection.text.trim() !== ''
                        ? `${key}: ${texts[currentLanguage].formatApplied}`
                        : `${key}: ${texts[currentLanguage].formatApplied} (次回入力用)`;
                    showMessage(message, 'success');

                    // 書式適用後にカーソル位置を復元
                    await restoreCursorPosition();

                } catch (error) {
                    console.error('書式適用エラー:', error);
                    console.error('Error details:', error.debugInfo);
                    showMessage('書式の適用に失敗しました', 'error');
                }
            }).catch(error => {
                console.error('Word.run エラー:', error);
                showMessage('書式の適用に失敗しました', 'error');
            });
        }

        // 現在の書式をアドイン内で管理・表示
        async function updateCurrentFormatDisplay(format) {
            try {
                console.log('🎨 Updating current format display:', format);
                
                // 現在の書式をグローバル変数に保存
                currentFormat = format;
                
                // フォントサイズと行間を更新
                if (format.font.size) {
                    currentFontSize = format.font.size;
                    updateFontSizeDisplay();
                }
                if (format.paragraph.lineSpacing) {
                    currentLineSpacing = format.paragraph.lineSpacing;
                    updateLineSpacingDisplay();
                }
                
                // 現在の書式表示を更新
                displayCurrentFormat(format);
                
                console.log('✅ Current format display updated successfully');
                console.log('📊 Current format:', {
                    fontSize: currentFontSize,
                    lineSpacing: currentLineSpacing,
                    fontName: format.font.name,
                    alignment: format.paragraph.alignment
                });

            } catch (error) {
                console.error('❌ Failed to update current format display:', error);
            }
        }

// 選択変更時の処理
function onSelectionChanged() {
    console.log('Selection changed');
    try {
        updateCurrentFormat();
        
        // 連続モードが有効で、書式が保存されている場合
        if (continuousMode && continuousFormat) {
            applyContinuousFormat();
        }
    } catch (error) {
        console.error('Selection change error:', error);
    }
}

// 現在の書式を更新
function updateCurrentFormat() {
    console.log('updateCurrentFormat called');
    
    if (typeof Word === 'undefined') {
        console.error('Word API not available');
        return;
    }
    
    Word.run(async (context) => {
        try {
            console.log('Word.run started');
            const selection = context.document.getSelection();
            
            // 選択範囲を確認
            selection.load('text');
            await context.sync();
            
            console.log('Selected text:', selection.text);
            
            // テキストが選択されているかチェック
            if (!selection.text || selection.text.trim() === '') {
                console.log('No text selected');
                currentFormat = null;
                displayCurrentFormat(null);
                return;
            }
            
            const font = selection.font;
            const paragraph = selection.paragraphs.getFirst();
            
            // 書式情報を読み込み
            font.load('name, size, bold, italic, color, underline, highlightColor');
            paragraph.load('alignment, leftIndent, rightIndent, lineSpacing, spaceAfter, spaceBefore');
            
            await context.sync();
            
            console.log('Font info:', {
                name: font.name,
                size: font.size,
                bold: font.bold,
                italic: font.italic,
                color: font.color
            });
            
            // 書式情報を取得
            currentFormat = {
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
            
            // 現在のフォントサイズと行間を更新
            currentFontSize = font.size;
            currentLineSpacing = paragraph.lineSpacing;
            
            // 表示を更新
            updateFontSizeDisplay();
            updateContinuousDisplay();
            
            // 現在の書式を表示
            displayCurrentFormat(currentFormat);
            console.log('Format updated successfully');
            
        } catch (error) {
            console.error('書式取得エラー:', error);
            currentFormat = null;
            displayCurrentFormat(null);
        }
    }).catch(error => {
        console.error('Word.run エラー:', error);
        currentFormat = null;
        displayCurrentFormat(null);
    });
}

// 現在の書式を表示
function displayCurrentFormat(format) {
    const formatDisplay = document.getElementById('current-format-display');
    
    if (!format) {
        formatDisplay.innerHTML = `<p>${texts[currentLanguage].noSelectionText}</p>`;
        return;
    }
    
    const font = format.font;
    const paragraph = format.paragraph;
    
    // 配置の日本語表示
    const alignmentText = getAlignmentText(paragraph.alignment);
    
    const formatText = `
        <div class="format-info">
            <strong>${font.name}</strong> ${font.size}px<br>
            ${font.bold ? '太字' : ''} ${font.italic ? '斜体' : ''}<br>
            ${alignmentText} | 色: ${font.color}
        </div>
    `;
    
    formatDisplay.innerHTML = formatText;
}

// 配置の日本語表示を取得
function getAlignmentText(alignment) {
    const alignments = {
        'Left': currentLanguage === 'ja' ? '左揃え' : 'Left',
        'Center': currentLanguage === 'ja' ? '中央揃え' : 'Center',
        'Right': currentLanguage === 'ja' ? '右揃え' : 'Right',
        'Justified': currentLanguage === 'ja' ? '両端揃え' : 'Justified'
    };
    return alignments[alignment] || alignment;
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
    const savedFormatsList = document.getElementById('saved-formats-list');
    
    if (Object.keys(savedFormats).length === 0) {
        savedFormatsList.innerHTML = `<p>${texts[currentLanguage].noSavedFormatsText}</p>`;
        return;
    }
    
    let html = '';
    for (const [key, format] of Object.entries(savedFormats)) {
        const date = new Date(format.timestamp).toLocaleDateString();
        html += `
            <div class="format-item" data-key="${key}" tabindex="0">
                <div>
                    <div class="format-key">${key}</div>
                    <div class="format-preview">${format.font.name} ${format.font.size}px - ${getAlignmentText(format.paragraph.alignment)} (${date})</div>
                </div>
                <button class="format-remove" data-key="${key}">×</button>
            </div>
        `;
    }
    
    savedFormatsList.innerHTML = html;
    
    // 削除ボタンのイベントリスナーを追加
    const removeButtons = savedFormatsList.querySelectorAll('.format-remove');
    removeButtons.forEach(button => {
        button.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const key = button.dataset.key;
            console.log('Delete button clicked for key:', key);
            removeFormat(key);
        });
        
        button.addEventListener('mousedown', (e) => {
            e.stopPropagation();
        });
    });
    
    // 書式項目のイベントリスナーを追加
    const formatItems = savedFormatsList.querySelectorAll('.format-item');
    formatItems.forEach(item => {
        item.addEventListener('mouseenter', async (e) => {
            console.log('🖱️ Format item mouseenter');
            e.preventDefault();
            await saveCursorPosition(); // カーソル位置を保存
            item.focus();
        });
        
        item.addEventListener('mouseleave', async (e) => {
            console.log('🖱️ Format item mouseleave');
            await restoreCursorPosition(); // カーソル位置を復元
        });
        
        item.addEventListener('keydown', (e) => {
            if (e.key !== 'Tab' && e.key !== 'Shift' && e.key !== 'Control' && 
                e.key !== 'Alt' && e.key !== 'Meta' && e.key !== 'CapsLock' &&
                e.key !== 'Enter' && e.key !== 'Escape' && e.key !== 'ArrowUp' &&
                e.key !== 'ArrowDown' && e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') {
                e.preventDefault();
                e.stopPropagation();
                const key = e.key.toLowerCase();
                loadFormat(key);
            }
        });
    });
    
    // スクロールビュー全体のイベントリスナーを追加
    const savedFormatsSection = document.getElementById('saved-formats-section');
    const instructionText = document.getElementById('saved-formats-instruction');
    
    if (savedFormatsSection && instructionText) {
        savedFormatsSection.addEventListener('mouseenter', () => {
            instructionText.style.opacity = '1';
        });
        
        savedFormatsSection.addEventListener('mouseleave', () => {
            instructionText.style.opacity = '0.7';
        });
    }
}

// 書式の削除
function removeFormat(key) {
    if (confirm(`書式 "${key}" を削除しますか？`)) {
        delete savedFormats[key];
        localStorage.setItem('savedFormats', JSON.stringify(savedFormats));
        updateSavedFormatsList();
        showMessage(`書式 "${key}" を削除しました`, 'success');
    }
}

// メッセージを表示
function showMessage(message, type) {
    // 既存のメッセージを削除
    const existingMessage = document.querySelector('.status-message');
    if (existingMessage) {
        existingMessage.remove();
    }
    
    // 新しいメッセージを作成
    const messageDiv = document.createElement('div');
    messageDiv.className = `status-message status-${type}`;
    messageDiv.textContent = message;
    
    // メッセージを表示
    document.body.appendChild(messageDiv);
    
    // 3秒後にメッセージを削除
    setTimeout(() => {
        if (messageDiv.parentNode) {
            messageDiv.remove();
        }
    }, 3000);
}

// 疑似クリックイベントの設定
function setupSyntheticClick() {
    console.log('=== setupSyntheticClick called ===');
    try {
        // 位置0,0での疑似クリックイベントを作成
        const syntheticClickEvent = new MouseEvent('click', {
            bubbles: true,
            cancelable: true,
            view: window,
            clientX: 0,
            clientY: 0,
            screenX: 0,
            screenY: 0,
            button: 0,
            buttons: 1,
            ctrlKey: false,
            shiftKey: false,
            altKey: false,
            metaKey: false
        });
        
        // 疑似クリックイベントを発火
        document.dispatchEvent(syntheticClickEvent);
        
        console.log('✅ Synthetic click event dispatched at position (0,0)');
        console.log('Event details:', {
            type: syntheticClickEvent.type,
            bubbles: syntheticClickEvent.bubbles,
            cancelable: syntheticClickEvent.cancelable
        });
    } catch (error) {
        console.error('❌ Synthetic click error:', error);
        console.error('Error stack:', error.stack);
    }
}

// Word APIの可用性チェック
function checkWordAPIAvailability() {
    console.log('=== Word API Availability Check ===');
    console.log('Check started at:', new Date().toISOString());
    
    // 1. Office.jsの読み込み確認
    if (typeof Office === 'undefined') {
        console.error('❌ Office.js is not loaded');
        showMessage('Office.jsが読み込まれていません', 'error');
        return false;
    }
    console.log('✅ Office.js is loaded');
    
    // 2. Office.contextの確認
    if (!Office.context) {
        console.error('❌ Office.context is not available');
        showMessage('Office.contextが利用できません', 'error');
        return false;
    }
    console.log('✅ Office.context is available');
    
    // 3. Word APIの確認
    if (typeof Word === 'undefined') {
        console.error('❌ Word API is not available');
        showMessage('Word APIが利用できません', 'error');
        return false;
    }
    console.log('✅ Word API is available');
    
    // 4. Office.context.documentの確認
    if (!Office.context.document) {
        console.error('❌ Office.context.document is not available');
        showMessage('Office.context.documentが利用できません', 'error');
        return false;
    }
    console.log('✅ Office.context.document is available');
    
    // 5. ホストアプリケーションの確認
    console.log('Host application:', Office.context.host);
    if (Office.context.host !== Office.HostType.Word) {
        console.warn('⚠️ Not running in Word host:', Office.context.host);
        showMessage('Word以外のアプリケーションで実行されています', 'error');
        return false;
    }
    console.log('✅ Running in Word host');
    
    // 6. プラットフォーム情報の確認
    console.log('Platform:', Office.context.platform);
    console.log('Office version:', Office.context.requirements);
    
    // 7. 基本的なWord API機能のテスト
    try {
        Word.run(async (context) => {
            const document = context.document;
            document.load('body');
            await context.sync();
            console.log('✅ Basic Word API test passed');
            console.log('Document body length:', document.body.text ? document.body.text.length : 0);
        }).catch(error => {
            console.error('❌ Basic Word API test failed:', error);
            showMessage('Word APIの基本テストに失敗しました', 'error');
        });
    } catch (error) {
        console.error('❌ Word API test error:', error);
        showMessage('Word APIテストでエラーが発生しました', 'error');
    }
    
    console.log('=== Word API Availability Check Complete ===');
    console.log('Check completed at:', new Date().toISOString());
    return true;
}

// フォントサイズ調整
function adjustFontSize(key) {
    const step = 1;
    if (key === '+' || key === '=') {
        currentFontSize += step;
    } else if (key === '-') {
        currentFontSize = Math.max(1, currentFontSize - step);
    } else {
        return;
    }
    
    updateFontSizeDisplay();
    applyCurrentFormat();
}

// 連続モード切り替え
function toggleContinuousMode() {
    continuousMode = !continuousMode;
    updateContinuousDisplay();
    
    const t = texts[currentLanguage];
    const message = continuousMode ? t.continuousModeEnabled : t.continuousModeDisabled;
    showMessage(message, 'success');
    
    console.log('🔄 Continuous mode:', continuousMode ? 'ON' : 'OFF');
}

// フォントサイズ表示更新
function updateFontSizeDisplay() {
    const display = document.getElementById('font-size-display');
    if (display) {
        display.textContent = `${currentFontSize}px`;
    }
}

// 連続モード表示更新
function updateContinuousDisplay() {
    const display = document.getElementById('continuous-display');
    if (display) {
        const t = texts[currentLanguage];
        if (continuousMode && continuousFormat) {
            // 保持している書式のキーを表示（最初の5文字）
            const formatKey = continuousFormat.key || 'FORMAT';
            display.textContent = formatKey.substring(0, 5);
        } else {
            display.textContent = t.continuousModeOff;
        }
    }
}

// 連続適用用の書式を設定（既存の保存された書式から取得）
function setContinuousFormat(key) {
    if (!savedFormats[key]) {
        showMessage(texts[currentLanguage].formatNotFound, 'error');
        return;
    }

    try {
        continuousFormat = {
            ...savedFormats[key],
            key: key,
            timestamp: new Date().toISOString()
        };

        const t = texts[currentLanguage];
        showMessage(`${key}: 連続適用用書式を設定しました`, 'success');
        
        // 表示を更新
        updateContinuousDisplay();
        
        console.log('💾 Continuous format set from saved format:', continuousFormat);
    } catch (error) {
        console.error('連続書式設定エラー:', error);
        showMessage('連続書式の設定に失敗しました', 'error');
    }
}

// 連続書式を適用
function applyContinuousFormat() {
    if (!continuousFormat) return;

    Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();

            // テキストが選択されている場合のみ適用
            if (selection.text && selection.text.trim() !== '') {
                console.log('🎨 Applying continuous format to:', selection.text);
                
                const font = selection.font;
                const paragraph = selection.paragraphs.getFirst();

                // フォント書式を適用
                if (continuousFormat.font.name) font.name = continuousFormat.font.name;
                if (continuousFormat.font.size) font.size = continuousFormat.font.size;
                if (continuousFormat.font.bold !== undefined) font.bold = continuousFormat.font.bold;
                if (continuousFormat.font.italic !== undefined) font.italic = continuousFormat.font.italic;
                if (continuousFormat.font.color) font.color = continuousFormat.font.color;
                if (continuousFormat.font.underline !== undefined) font.underline = continuousFormat.font.underline;
                if (continuousFormat.font.highlightColor) font.highlightColor = continuousFormat.font.highlightColor;

                // 段落書式を適用
                if (continuousFormat.paragraph.alignment) paragraph.alignment = continuousFormat.paragraph.alignment;
                if (continuousFormat.paragraph.leftIndent !== undefined) paragraph.leftIndent = continuousFormat.paragraph.leftIndent;
                if (continuousFormat.paragraph.rightIndent !== undefined) paragraph.rightIndent = continuousFormat.paragraph.rightIndent;
                if (continuousFormat.paragraph.lineSpacing !== undefined) paragraph.lineSpacing = continuousFormat.paragraph.lineSpacing;
                if (continuousFormat.paragraph.spaceAfter !== undefined) paragraph.spaceAfter = continuousFormat.paragraph.spaceAfter;
                if (continuousFormat.paragraph.spaceBefore !== undefined) paragraph.spaceBefore = continuousFormat.paragraph.spaceBefore;

                await context.sync();
                console.log('✅ Continuous format applied successfully');
            }
        } catch (error) {
            console.error('連続書式適用エラー:', error);
        }
    }).catch(error => {
        console.error('Word.run エラー:', error);
    });
}

        // 現在の書式を適用
        function applyCurrentFormat() {
            if (!currentFormat) return;

            Word.run(async (context) => {
                try {
                    const selection = context.document.getSelection();
                    const font = selection.font;
                    const paragraph = selection.paragraphs.getFirst();

                    // 現在の書式を更新
                    currentFormat.font.size = currentFontSize;
                    currentFormat.paragraph.lineSpacing = currentLineSpacing;

                    console.log('🎨 Applying current format:', {
                        fontSize: currentFontSize,
                        lineSpacing: currentLineSpacing
                    });

                    // 書式を適用（安全な方法）
                    if (currentFormat.font.name) font.name = currentFormat.font.name;
                    if (currentFormat.font.size) font.size = currentFormat.font.size;
                    if (currentFormat.font.bold !== undefined) font.bold = currentFormat.font.bold;
                    if (currentFormat.font.italic !== undefined) font.italic = currentFormat.font.italic;
                    if (currentFormat.font.color) font.color = currentFormat.font.color;
                    if (currentFormat.font.underline !== undefined) font.underline = currentFormat.font.underline;
                    if (currentFormat.font.highlightColor) font.highlightColor = currentFormat.font.highlightColor;

                    if (currentFormat.paragraph.alignment) paragraph.alignment = currentFormat.paragraph.alignment;
                    if (currentFormat.paragraph.leftIndent !== undefined) paragraph.leftIndent = currentFormat.paragraph.leftIndent;
                    if (currentFormat.paragraph.rightIndent !== undefined) paragraph.rightIndent = currentFormat.paragraph.rightIndent;
                    if (currentFormat.paragraph.lineSpacing !== undefined) paragraph.lineSpacing = currentFormat.paragraph.lineSpacing;
                    if (currentFormat.paragraph.spaceAfter !== undefined) paragraph.spaceAfter = currentFormat.paragraph.spaceAfter;
                    if (currentFormat.paragraph.spaceBefore !== undefined) paragraph.spaceBefore = currentFormat.paragraph.spaceBefore;

                    await context.sync();
                    console.log('✅ Current format applied successfully');

                    // アドイン内の書式表示も更新
                    await updateCurrentFormatDisplay(currentFormat);

                } catch (error) {
                    console.error('書式適用エラー:', error);
                    console.error('Error details:', error.debugInfo);
                }
            }).catch(error => {
                console.error('Word.run エラー:', error);
            });
        }

// ホイールイベント処理
function handleFontWheel(event) {
    event.preventDefault();
    const delta = event.deltaY > 0 ? -1 : 1;
    currentFontSize = Math.max(1, currentFontSize + delta);
    updateFontSizeDisplay();
    applyCurrentFormat();
}

function handleLineSpacingWheel(event) {
    event.preventDefault();
    const delta = event.deltaY > 0 ? -0.5 : 0.5;
    currentLineSpacing = Math.max(0.5, currentLineSpacing + delta);
    updateLineSpacingDisplay();
    applyCurrentFormat();
}

// 幅切り替え
function toggleWidth() {
    isWideMode = !isWideMode;
    const app = document.getElementById('app');
    const button = document.getElementById('width-toggle');
    
    if (isWideMode) {
        app.classList.remove('narrow');
        app.classList.add('wide');
        button.textContent = '幅: 300px';
    } else {
        app.classList.remove('wide');
        app.classList.add('narrow');
        button.textContent = '幅: 100px';
    }
}

// グローバル関数として公開
window.removeFormat = removeFormat;

// デバッグ用: 手動初期化
window.manualInit = function() {
    console.log('Manual initialization triggered');
    window.appInitialized = false;
    initializeApp();
};