if (typeof module !== 'undefined') module.exports = {
  SpreadsheetApp: {
    getActiveSpreadsheet: jest.fn(() => ({
      getSheetByName: jest.fn()
    })),
    getActiveUser: jest.fn(() => ({
      getEmail: jest.fn(() => 'test@example.com')
    }))
  },
  Session: {
    getActiveUser: jest.fn(() => ({
      getEmail: jest.fn(() => 'test@example.com')
    }))
  }
};
