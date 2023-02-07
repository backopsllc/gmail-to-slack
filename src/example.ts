/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
// Optional Types
const isDone = false;
const height = 6;
const bob = 'bob';
const list1: number[] = [1, 2, 3];
const list2: Array<number> = [1, 2, 3];
enum Color {
  Red,
  Green,
  Blue,
}
const c: Color = Color.Green;
let notSure: any = 4;
notSure = 'maybe a string instead';
notSure = false; // okay, definitely a boolean
function showMessage(data: string): void {
  // Void
  console.log(data);
}
showMessage('hello');

// Classes
class Hamburger {
  constructor() {
    // This is the constructor.
  }
  listToppings() {
    // This is a method.
  }
}

// Template strings
const name = 'Sam';
const age = 42;
console.log(`hello my name is ${name}, and I am ${age} years old`);

// Rest arguments
const buildName = (first: string, ...rest: string[]): string =>
  first + ' ' + rest.join(' ');
buildName('First', 'Second', 'Third');

// Spread operator (array)
const cde = ['c', 'd', 'e'];
const scale = ['a', 'b', ...cde, 'f', 'g']; // ['a', 'b', 'c', 'd', 'e', 'f', 'g']

// Spread operator (map)
const mapABC = {a: 5, b: 6, c: 3};
const mapABCD = {...mapABC, d: 7}; // { a: 5, b: 6, c: 3, d: 7 }

// Destructure map
const jane = {firstName: 'Jane', lastName: 'Doe'};
const john = {firstName: 'John', lastName: 'Doe', middleName: 'Smith'};
function sayName({firstName = '', lastName = '', middleName = 'N/A'}): void {
  console.log(`Hello ${firstName} ${middleName} ${lastName}`);
}
sayName(jane); // -> Hello Jane N/A Doe
sayName(john); // -> Hello John Smith Doe

// Export (The export keyword is ignored)
export const pi = 3.141592;

// Google Apps Script Services
function createDocument(): string {
  const doc = DocumentApp.create('Hello, world!');
  doc
    .getBody()
    .appendParagraph('This document was created by Google Apps Script.');
  return doc.getId();
}

// Decorators
function Override(label: string) {
  return function (target: any, key: string): void {
    Object.defineProperty(target, key, {
      configurable: false,
      get: () => label,
      set: target => {
        key = target;
      },
    });
  };
}
class Test {
  @Override('test') // invokes Override, which returns the decorator
  name = 'pat';
}
const t = new Test();
console.log(t.name); // 'test'

export {
  isDone,
  height,
  bob,
  list1,
  list2,
  Color,
  notSure,
  showMessage,
  Hamburger,
  buildName,
  scale,
  mapABCD,
  jane,
  john,
  sayName,
  Test,
  createDocument,
};
