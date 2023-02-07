module.exports = {
  presets: [
    [
      // ES features necessary for user's Node version
      '@babel/preset-env',
      {
        targets: {
          node: 'current',
        },
      },
    ],
    ['@babel/preset-typescript'],
  ],
  plugins: [
    // Polyfills the runtime needed for async/await, generators, and friends
    // https://babeljs.io/docs/en/babel-plugin-transform-runtime
    ['@babel/plugin-transform-runtime'],
    // https://babeljs.io/docs/en/babel-plugin-proposal-decorators
    ['@babel/plugin-proposal-decorators', {version: 'legacy'}],
    '@babel/plugin-proposal-class-properties',
  ],
};
