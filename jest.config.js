module.exports = {
  testEnvironment: 'node',
  testMatch: ['**/__tests__/**/*.js', '**/*.test.js'],
  collectCoverageFrom: [
    'lib/**/*.js',
    'gdocs-me-up.js'
  ],
  coveragePathIgnorePatterns: [
    '/node_modules/'
  ]
};
