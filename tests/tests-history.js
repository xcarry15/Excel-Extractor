// tests/tests-history.js
import {
  loadHistories,
  saveHistory,
  clearHistories,
  getHistoryDisplayText
} from '../src/services/history.js';
import { HISTORY_KEY } from '../src/constants.js';

QUnit.module('History Service', function() {

  QUnit.test('loadHistories returns empty array when no history', function(assert) {
    clearHistories();
    const histories = loadHistories();
    assert.ok(Array.isArray(histories), 'returns array');
    assert.equal(histories.length, 0, 'empty when no history');
  });

  QUnit.test('saveHistory adds new history', function(assert) {
    clearHistories();
    saveHistory('Config1', ['A', 'B']);
    const histories = loadHistories();
    assert.equal(histories.length, 1, 'has 1 history');
    assert.equal(histories[0].name, 'Config1', 'name is correct');
    assert.deepEqual(histories[0].columns, ['A', 'B'], 'columns correct');
  });

  QUnit.test('saveHistory limits to 20 items', function(assert) {
    clearHistories();
    for (let i = 0; i < 25; i++) {
      saveHistory(`Config${i}`, [`Col${i}`]);
    }
    const histories = loadHistories();
    assert.equal(histories.length, 20, 'limited to 20 items');
    assert.equal(histories[0].name, 'Config24', 'newest first');
  });

  QUnit.test('saveHistory deduplicates same columns', function(assert) {
    clearHistories();
    saveHistory('Config1', ['A', 'B']);
    saveHistory('Config2', ['A', 'B']); // same columns
    saveHistory('Config3', ['C']);
    const histories = loadHistories();
    assert.equal(histories.length, 2, 'deduplicated');
    assert.equal(histories[0].name, 'Config3', 'newest first');
  });

  QUnit.test('clearHistories removes all history', function(assert) {
    clearHistories();
    saveHistory('Config1', ['A']);
    clearHistories();
    const histories = loadHistories();
    assert.equal(histories.length, 0, 'all cleared');
  });

  QUnit.test('getHistoryDisplayText formats correctly', function(assert) {
    clearHistories();
    saveHistory('MyConfig', ['Col1', 'Col2', 'Col3']);
    const histories = loadHistories();
    const text = getHistoryDisplayText(histories[0]);
    assert.equal(text, 'MyConfig · Col1, Col2, Col3', 'display text correct');
  });

});
