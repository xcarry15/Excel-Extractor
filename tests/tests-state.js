// tests/tests-state.js
import {
  getState,
  setState,
  resetState,
  rebuildHeaderIndexMap,
  updateSelectedDerived,
  addSelected,
  removeSelectedByIndex,
  clearAllSelected,
  setParseResult,
  applyHistoryConfig,
  subscribe
} from '../src/state.js';

QUnit.module('State Management', function() {

  QUnit.test('getState returns initial state', function(assert) {
    resetState();
    const state = getState();
    assert.ok(Array.isArray(state.dataRows), 'dataRows is array');
    assert.ok(Array.isArray(state.headers), 'headers is array');
    assert.ok(Array.isArray(state.selected), 'selected is array');
    assert.ok(Array.isArray(state.selectedWithIndex), 'selectedWithIndex is array');
    assert.ok(Array.isArray(state.skippedItems), 'skippedItems is array');
    assert.equal(state.filename, '', 'filename is empty');
  });

  QUnit.test('setState updates state correctly', function(assert) {
    resetState();
    setState({ filename: 'test.xlsx', headers: ['A', 'B', 'C'] });
    const state = getState();
    assert.equal(state.filename, 'test.xlsx', 'filename updated');
    assert.equal(state.headers.length, 3, 'headers updated');
  });

  QUnit.test('rebuildHeaderIndexMap creates correct mapping', function(assert) {
    resetState();
    setState({ headers: ['X', 'Y', 'Z'] });
    rebuildHeaderIndexMap();
    const state = getState();
    assert.equal(state.headerIndexMap.get('X'), 0, 'X maps to 0');
    assert.equal(state.headerIndexMap.get('Y'), 1, 'Y maps to 1');
    assert.equal(state.headerIndexMap.get('Z'), 2, 'Z maps to 2');
  });

  QUnit.test('updateSelectedDerived calculates indices correctly', function(assert) {
    resetState();
    setState({ headers: ['A', 'B', 'C', 'D'] });
    rebuildHeaderIndexMap();
    setState({ selected: ['C', 'A', 'D'] });
    updateSelectedDerived();
    const state = getState();
    assert.deepEqual(state.selectedIdx, [2, 0, 3], 'selectedIdx matches');
  });

  QUnit.test('subscribe and notify works', function(assert) {
    resetState();
    let notified = false;
    let receivedState = null;
    const unsubscribe = subscribe((state) => {
      notified = true;
      receivedState = state;
    });
    setState({ filename: 'test' });
    assert.ok(notified, 'subscriber was notified');
    assert.equal(receivedState.filename, 'test', 'received correct state');
    unsubscribe();
    notified = false;
    setState({ filename: 'test2' });
    assert.ok(!notified, 'unsubscribed does not receive notifications');
  });

  QUnit.test('addSelected adds columns correctly', function(assert) {
    resetState();
    setState({ headers: ['Col1', 'Col2', 'Col3', 'Col1'] });
    rebuildHeaderIndexMap();
    addSelected(['Col1', 'Col2']);
    const state = getState();
    assert.equal(state.selected.length, 2, 'selected has 2 items');
    assert.equal(state.selected[0], 'Col1', 'first selected is Col1');
    assert.equal(state.selectedWithIndex[0].originalIndex, 0, 'Col1 original index is 0');
    assert.equal(state.selectedWithIndex[1].originalIndex, 1, 'Col2 original index is 1');
  });

  QUnit.test('addSelected with firstOnly=false adds duplicate columns', function(assert) {
    resetState();
    setState({ headers: ['Col1', 'Col2', 'Col1', 'Col3'] });
    rebuildHeaderIndexMap();
    addSelected(['Col1'], false);
    const state = getState();
    assert.equal(state.selected.length, 2, 'selected has 2 items (duplicates)');
    assert.equal(state.selectedWithIndex[0].originalIndex, 0, 'first Col1 index is 0');
    assert.equal(state.selectedWithIndex[1].originalIndex, 2, 'second Col1 index is 2');
  });

  QUnit.test('addSelected skips non-existent columns', function(assert) {
    resetState();
    setState({ headers: ['Col1', 'Col2'] });
    rebuildHeaderIndexMap();
    addSelected(['Col1', 'NonExistent', 'Col2']);
    const state = getState();
    assert.equal(state.selected.length, 2, 'selected has 2 items');
    assert.equal(state.skippedItems.length, 1, 'one item skipped');
    assert.equal(state.skippedItems[0].name, 'NonExistent', 'skipped item is NonExistent');
  });

  QUnit.test('removeSelectedByIndex removes correct item', function(assert) {
    resetState();
    setState({ headers: ['A', 'B', 'C'] });
    rebuildHeaderIndexMap();
    addSelected(['A', 'B', 'C']);
    removeSelectedByIndex(1);
    const state = getState();
    assert.equal(state.selected.length, 2, 'selected has 2 items');
    assert.deepEqual(state.selected, ['A', 'C'], 'B was removed');
  });

  QUnit.test('clearAllSelected clears all items', function(assert) {
    resetState();
    setState({ headers: ['A', 'B', 'C'] });
    rebuildHeaderIndexMap();
    addSelected(['A', 'B']);
    clearAllSelected();
    const state = getState();
    assert.equal(state.selected.length, 0, 'selected is empty');
    assert.equal(state.selectedWithIndex.length, 0, 'selectedWithIndex is empty');
  });

  QUnit.test('setParseResult resets and sets state', function(assert) {
    resetState();
    setState({ selected: ['A'] }); // 先设置一些值
    setParseResult({
      headers: ['X', 'Y'],
      dataRows: [[1, 2], [3, 4]],
      filename: 'test.xlsx'
    });
    const state = getState();
    assert.equal(state.headers.length, 2, 'headers set correctly');
    assert.equal(state.dataRows.length, 2, 'dataRows set correctly');
    assert.equal(state.filename, 'test.xlsx', 'filename set correctly');
    assert.equal(state.selected.length, 0, 'selected was cleared');
    assert.equal(state.skippedItems.length, 0, 'skippedItems was cleared');
  });

  QUnit.test('applyHistoryConfig restores configuration', function(assert) {
    resetState();
    setState({
      headers: ['A', 'B', 'C', 'D', 'A'],
      selected: [],
      selectedWithIndex: []
    });
    rebuildHeaderIndexMap();
    applyHistoryConfig(['C', 'A', 'D']);
    const state = getState();
    assert.equal(state.selected.length, 3, '3 columns restored');
    assert.equal(state.selected[0], 'C', 'first is C');
    assert.equal(state.selectedIdx[0], 2, 'C is at index 2');
  });

});
