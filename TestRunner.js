/**
 * ===================================================================
 * TEST RUNNER SERVICE
 * ===================================================================
 * Comprehensive testing framework for the fleet resource allocator
 */

/**
 * Test result object
 * @typedef {Object} TestResult
 * @property {string} name - Test name
 * @property {boolean} passed - Whether test passed
 * @property {string} message - Test message
 * @property {number} duration - Test duration in ms
 * @property {Error} [error] - Error if test failed
 */

/**
 * Test suite class
 */
class TestSuite {
  constructor(name) {
    this.name = name;
    this.tests = [];
    this.beforeEach = null;
    this.afterEach = null;
    this.logger = createLogger(`TestSuite:${name}`);
  }
  
  /**
   * Add a test to the suite
   * @param {string} name - Test name
   * @param {Function} testFn - Test function
   */
  test(name, testFn) {
    this.tests.push({ name, testFn });
  }
  
  /**
   * Set function to run before each test
   * @param {Function} fn - Setup function
   */
  setup(fn) {
    this.beforeEach = fn;
  }
  
  /**
   * Set function to run after each test
   * @param {Function} fn - Teardown function
   */
  teardown(fn) {
    this.afterEach = fn;
  }
  
  /**
   * Run all tests in the suite
   * @return {TestResult[]} Test results
   */
  run() {
    const results = [];
    
    this.logger.info(`Running test suite: ${this.name}`);
    
    for (const test of this.tests) {
      const startTime = new Date().getTime();
      let passed = false;
      let message = '';
      let error = null;
      
      try {
        // Run setup
        if (this.beforeEach) {
          this.beforeEach();
        }
        
        // Run test
        test.testFn();
        passed = true;
        message = 'Test passed';
        
      } catch (e) {
        passed = false;
        message = e.message;
        error = e;
        this.logger.error(`Test failed: ${test.name}`, { error: e.message });
      } finally {
        // Run teardown
        if (this.afterEach) {
          try {
            this.afterEach();
          } catch (e) {
            this.logger.error('Teardown failed', { error: e.message });
          }
        }
      }
      
      const duration = new Date().getTime() - startTime;
      
      results.push({
        name: test.name,
        passed,
        message,
        duration,
        error
      });
    }
    
    return results;
  }
}

/**
 * Test runner class
 */
class TestRunner {
  constructor() {
    this.suites = [];
    this.logger = createLogger('TestRunner');
  }
  
  /**
   * Add a test suite
   * @param {TestSuite} suite - Test suite to add
   */
  addSuite(suite) {
    this.suites.push(suite);
  }
  
  /**
   * Run all test suites
   * @return {Object} Test results summary
   */
  runAll() {
    const startTime = new Date().getTime();
    const results = {
      suites: [],
      totalTests: 0,
      passed: 0,
      failed: 0,
      duration: 0
    };
    
    this.logger.info('Starting test run');
    
    for (const suite of this.suites) {
      const suiteResults = suite.run();
      
      const suiteData = {
        name: suite.name,
        tests: suiteResults,
        passed: suiteResults.filter(r => r.passed).length,
        failed: suiteResults.filter(r => !r.passed).length
      };
      
      results.suites.push(suiteData);
      results.totalTests += suiteResults.length;
      results.passed += suiteData.passed;
      results.failed += suiteData.failed;
    }
    
    results.duration = new Date().getTime() - startTime;
    
    this.logger.info('Test run complete', {
      totalTests: results.totalTests,
      passed: results.passed,
      failed: results.failed,
      duration: results.duration
    });
    
    return results;
  }
  
  /**
   * Generate HTML report of test results
   * @param {Object} results - Test results from runAll
   * @return {string} HTML report
   */
  generateReport(results) {
    let html = `
      <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .summary { background: #f0f0f0; padding: 15px; margin-bottom: 20px; border-radius: 5px; }
          .suite { margin-bottom: 30px; }
          .suite-header { background: #e0e0e0; padding: 10px; font-weight: bold; }
          .test { padding: 8px 20px; border-bottom: 1px solid #eee; }
          .test.passed { background: #e8f5e9; }
          .test.failed { background: #ffebee; }
          .duration { color: #666; font-size: 0.9em; }
          .error { color: #d32f2f; margin-top: 5px; font-size: 0.9em; }
        </style>
      </head>
      <body>
        <h1>Test Results</h1>
        <div class="summary">
          <h2>Summary</h2>
          <p>Total Tests: ${results.totalTests}</p>
          <p>Passed: ${results.passed}</p>
          <p>Failed: ${results.failed}</p>
          <p>Duration: ${results.duration}ms</p>
        </div>
    `;
    
    for (const suite of results.suites) {
      html += `
        <div class="suite">
          <div class="suite-header">${suite.name} (${suite.passed}/${suite.tests.length} passed)</div>
      `;
      
      for (const test of suite.tests) {
        html += `
          <div class="test ${test.passed ? 'passed' : 'failed'}">
            ${test.passed ? '✓' : '✗'} ${test.name}
            <span class="duration">(${test.duration}ms)</span>
            ${test.error ? `<div class="error">${test.error.message}</div>` : ''}
          </div>
        `;
      }
      
      html += '</div>';
    }
    
    html += '</body></html>';
    
    return html;
  }
}

/**
 * Assertion utilities
 */
const assert = {
  /**
   * Assert that a value is truthy
   * @param {*} value - Value to test
   * @param {string} [message] - Error message
   */
  ok(value, message) {
    if (!value) {
      throw new Error(message || `Expected truthy value, got ${value}`);
    }
  },
  
  /**
   * Assert that values are equal
   * @param {*} actual - Actual value
   * @param {*} expected - Expected value
   * @param {string} [message] - Error message
   */
  equal(actual, expected, message) {
    if (actual !== expected) {
      throw new Error(message || `Expected ${expected}, got ${actual}`);
    }
  },
  
  /**
   * Assert that values are deeply equal
   * @param {*} actual - Actual value
   * @param {*} expected - Expected value
   * @param {string} [message] - Error message
   */
  deepEqual(actual, expected, message) {
    if (JSON.stringify(actual) !== JSON.stringify(expected)) {
      throw new Error(message || `Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
    }
  },
  
  /**
   * Assert that a function throws an error
   * @param {Function} fn - Function to test
   * @param {string} [message] - Error message
   */
  throws(fn, message) {
    let threw = false;
    try {
      fn();
    } catch (e) {
      threw = true;
    }
    if (!threw) {
      throw new Error(message || 'Expected function to throw');
    }
  }
};

/**
 * Run all tests and display results
 */
function runAllTests() {
  const runner = new TestRunner();
  
  // Add test suites
  runner.addSuite(createUtilityTests());
  runner.addSuite(createAllocationTests());
  runner.addSuite(createFormTests());
  runner.addSuite(createEmailTests());
  
  // Run tests
  const results = runner.runAll();
  
  // Show results
  const html = HtmlService.createHtmlOutput(runner.generateReport(results))
    .setWidth(800)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Test Results');
}

/**
 * Create utility function tests
 * @return {TestSuite} Test suite
 */
function createUtilityTests() {
  const suite = new TestSuite('Utility Functions');
  
  suite.test('extractFileId - extracts from URL', () => {
    const id = extractFileId('https://docs.google.com/spreadsheets/d/1abc123def456/edit');
    assert.equal(id, '1abc123def456');
  });
  
  suite.test('extractFileId - handles direct ID', () => {
    const id = extractFileId('1abc123def456ghi789');
    assert.equal(id, '1abc123def456ghi789');
  });
  
  suite.test('extractFileId - throws on invalid input', () => {
    assert.throws(() => extractFileId('invalid'));
  });
  
  suite.test('formatDate - formats date correctly', () => {
    const date = new Date('2024-12-25');
    const formatted = formatDate(date);
    assert.equal(formatted, '12/25/2024');
  });
  
  suite.test('getUSWeekNumber - calculates week correctly', () => {
    const week1 = getUSWeekNumber(new Date('2024-01-01'));
    assert.equal(week1, 1);
    
    const week52 = getUSWeekNumber(new Date('2024-12-25'));
    assert.equal(week52, 52);
  });
  
  return suite;
}

/**
 * Create allocation logic tests
 * @return {TestSuite} Test suite
 */
function createAllocationTests() {
  const suite = new TestSuite('Allocation Logic');
  
  suite.test('getVanType - maps service types correctly', () => {
    assert.equal(getVanType('Standard Parcel - Extra Large Van - US'), 'Extra Large');
    assert.equal(getVanType('Standard Parcel - Large Van'), 'Large');
    assert.equal(getVanType('Standard Parcel Step Van - US'), 'Step Van');
  });
  
  suite.test('getVanType - handles nursery routes', () => {
    assert.equal(getVanType('Nursery Route Level 1'), 'Large');
    assert.equal(getVanType('Nursery Route Level 2'), 'Large');
  });
  
  return suite;
}

/**
 * Create form functionality tests
 * @return {TestSuite} Test suite
 */
function createFormTests() {
  const suite = new TestSuite('Form Functions');
  
  suite.test('normalizeReportingTime - removes End of Day suffix', () => {
    const normalized = normalizeReportingTime('9:40 PM (End of Day)');
    assert.equal(normalized, '9:40 PM');
  });
  
  suite.test('formatTimeString - handles Date objects', () => {
    const date = new Date('2024-01-01 18:30:00');
    const formatted = formatTimeString(date);
    assert.ok(formatted.includes('PM'));
  });
  
  return suite;
}

/**
 * Create email service tests
 * @return {TestSuite} Test suite
 */
function createEmailTests() {
  const suite = new TestSuite('Email Service');
  
  suite.test('escapeHtml - escapes special characters', () => {
    const escaped = escapeHtml('<script>alert("test")</script>');
    assert.ok(!escaped.includes('<script>'));
    assert.ok(escaped.includes('&lt;script&gt;'));
  });
  
  suite.test('formatTimeWithTimezone - adds timezone', () => {
    const formatted = formatTimeWithTimezone('18:30');
    assert.ok(formatted.includes('PM'));
    assert.ok(formatted.includes('EST') || formatted.includes('EDT'));
  });
  
  return suite;
}

/**
 * Developer menu handler for running tests
 */
function testAllocationLogic() {
  runAllTests();
}

/**
 * Developer menu handler for form tests
 */
function testFormFunctionality() {
  const runner = new TestRunner();
  runner.addSuite(createFormTests());
  const results = runner.runAll();
  
  const html = HtmlService.createHtmlOutput(runner.generateReport(results))
    .setWidth(600)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Form Test Results');
}