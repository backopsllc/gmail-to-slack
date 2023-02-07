import {greeter, testGreeter} from '../hello';

describe('test.ts test', () => {
  beforeAll(() => {
    Logger.log = jest.fn().mockImplementation(msg => {
      return console.log(msg);
    });
    jest.spyOn(Logger, 'log');
  });
  test('greeter', () => {
    const person = 'World';
    const expected = greeter(person);
    expect(expected).toBe('Hello, World!');
  });
  test('testGreeter', () => {
    const expected = testGreeter();
    expect(Logger.log).toBeCalled();
    expect(expected).toBe('Hello, Grant!');
  });
});
