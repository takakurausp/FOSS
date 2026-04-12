// DecisionHistory.test.js
// GAS環境では実行されません / Not executed in GAS environment

if (typeof jest !== 'undefined') {

  const mockSpreadsheetApp = {
    getActiveSpreadsheet: jest.fn(() => ({
      getSheetByName: jest.fn()
    })),
    getActiveUser: jest.fn(() => ({
      getEmail: jest.fn(() => 'test@example.com')
    }))
  };

  const mockSession = {
    getActiveUser: jest.fn(() => ({
      getEmail: jest.fn(() => 'test@example.com')
    }))
  };

  global.SpreadsheetApp = mockSpreadsheetApp;
  global.Session = mockSession;

  const logDecisionHistory = require('../DecisionHistoryTrigger');

  describe('Decision History Logging System', () => {
    const mockAppend = jest.fn();

    beforeEach(() => {
      jest.clearAllMocks();
    });

    test('should correctly log decision changes', () => {
      const mockEvent = {
        range: {
          getRow: jest.fn(() => 2),
          getSheet: jest.fn(() => ({
            getName: jest.fn(() => 'Decisions'),
            getRange: jest.fn(() => ({
              getValues: jest.fn(() => [['MS-001', 'Accept', 'yes', 'no']])
            })),
            getLastColumn: jest.fn(() => 4)
          }))
        },
        oldValue: 'Minor Revision',
        value: 'Accept',
        source: {
          getSheetByName: jest.fn(() => ({
            appendRow: mockAppend
          }))
        }
      };

      logDecisionHistory(mockEvent);

      expect(mockAppend).toHaveBeenCalledWith([
        2,
        JSON.stringify({
          decision: 'Minor Revision',
          isAccepted: 'yes',
          resubmit: 'no'
        }),
        JSON.stringify({
          decision: 'Accept',
          isAccepted: 'yes',
          resubmit: 'no'
        }),
        'test@example.com',
        expect.any(Date)
      ]);
    });

    test('should skip logging for non-Decisions sheet edits', () => {
      const mockEvent = {
        range: {
          getRow: jest.fn(),
          getSheet: jest.fn(() => ({
            getName: jest.fn(() => 'OtherSheet')
          }))
        },
        source: {},
        oldValue: '',
        value: ''
      };

      logDecisionHistory(mockEvent);
      expect(mockAppend).not.toHaveBeenCalled();
    });
  });

}
