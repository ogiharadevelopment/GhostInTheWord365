// キャッシュバスター - ファイル更新を強制
console.log('Cache buster loaded at:', new Date().toISOString());

// バージョン情報
        const VERSION = '1.2.3.0';
console.log('Format Manager Add-in Version:', VERSION);

// キャッシュクリアの実行
if (typeof(Storage) !== "undefined") {
    // ローカルストレージのクリア（必要に応じて）
    // localStorage.clear();
    console.log('Local storage available');
} else {
    console.log('Local storage not available');
}

// ページの再読み込みを強制（開発時のみ）
// location.reload(true);
