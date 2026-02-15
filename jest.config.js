module.exports = {
  testEnvironment: 'node',
  testMatch: ['**/__tests__/**/*.js', '**/*.test.js'],
  testPathIgnorePatterns: [
    '/node_modules/',
    '/tests/visual/'  // Playwright tests, not Jest
  ],
  collectCoverageFrom: [
    'lib/**/*.js',
    'gdocs-me-up.js'
  ],
  coveragePathIgnorePatterns: [
    '/node_modules/'
  ]
};
