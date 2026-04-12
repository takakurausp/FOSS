// DECISION_HISTORYシートスキーマ定義
const DECISION_HISTORY_SCHEMA = {
  columns: [
    { name: 'decisionId', type: 'STRING', description: 'Decisionsシートの行IDと紐付け' },
    { name: 'previousState', type: 'JSON', description: '変更前の決定状態' },
    { name: 'newState', type: 'JSON', description: '変更後の決定状態' },
    { name: 'changedBy', type: 'STRING', description: '変更実行者のメールアドレス' },
    { name: 'timestamp', type: 'TIMESTAMP', description: '変更日時' }
  ],
  indexes: [
    { columns: ['decisionId'], unique: false },
    { columns: ['timestamp'], unique: false }
  ]
};