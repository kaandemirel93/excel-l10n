module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  testMatch: ['**/tests/**/*.test.ts'],
  moduleFileExtensions: ['ts', 'js', 'json'],
  transform: {
    '^.+\\.(ts|tsx)$': [
      'ts-jest',
      { tsconfig: { isolatedModules: true, module: 'CommonJS', target: 'ES2020' } }
    ],
  },
  moduleNameMapper: {
    // Map ESM-style internal imports with .js extension back to TS sources for ts-jest
    '^(\\.{1,2}/.*)\\.js$': '$1',
  },
};
