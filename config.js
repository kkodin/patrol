// 社内パトロール点検簿をつくるページの設定（GitHub に公開される。社員名・メール・共有リンクは書かないこと）
// サインインは「全体工程表」と同じ Entra のアプリ（リダイレクト URI に このページのアドレスを足してある）
window.PATROL_CONFIG = {
  clientId: "9a4c8447-3e6a-437b-85d0-4662ded35f42",   // アプリケーション (クライアント) ID（全体工程表と同じ）
  tenantId: "acd238cd-b6aa-4b5a-8056-44ed0177515f",   // ディレクトリ (テナント) ID
  folder: {       // 社内共有フォルダ\05_現場パトロール（原本の点検簿と データ\ が入っている）
    driveId: "b!fW3m73KhV0i7xf8c58hB1ha1ymrf1RZIq3ibCMmfIEMZJwRVBBcZSqMjvDbFdiXW",
    itemId: "",   // 空のときは持ち主の OneDrive から path で探す（見つかった ID を画面に出すので、ここへ貼る）
    path: "社内共有フォルダ/05_現場パトロール"
  }
};
