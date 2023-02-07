import {
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
  pi,
  Test,
  createDocument,
} from '../example';
import * as fs from 'fs';

class BodyMock {
  private _doc: DocumentMock;
  constructor(doc: DocumentMock) {
    this._doc = doc;
  }
  appendParagraph(text: string): void {
    fs.appendFile(this._doc.getName(), text, err => {
      if (err) {
        throw err;
      }
    });
  }
}
class DocumentMock {
  private _name: string;
  constructor(name: string) {
    this._name = name;
  }
  getBody(): BodyMock {
    return new BodyMock(this);
  }
  getId(): string {
    return '0123456789';
  }
  getName(): string {
    return this._name;
  }
}

describe('example.ts test', () => {
  test('variable', () => {
    expect(isDone).toBe(false);
    expect(height).toBe(6);
    expect(bob).toBe('bob');
    expect(list1[0]).toBe(1);
    expect(list1[1]).toBe(2);
    expect(list1[2]).toBe(3);
    expect(list2[0]).toBe(1);
    expect(list2[1]).toBe(2);
    expect(list2[2]).toBe(3);
    expect(Color).toHaveProperty('Red');
    expect(Color).toHaveProperty('Green');
    expect(Color).toHaveProperty('Blue');
    expect(Color.Red).toBe(0);
    expect(Color.Green).toBe(1);
    expect(Color.Blue).toBe(2);
    expect(notSure).toBe(false);
    expect(scale[0]).toBe('a');
    expect(scale[1]).toBe('b');
    expect(scale[2]).toBe('c');
    expect(scale[3]).toBe('d');
    expect(scale[4]).toBe('e');
    expect(scale[5]).toBe('f');
    expect(scale[6]).toBe('g');
    expect(mapABCD['a']).toBe(5);
    expect(mapABCD['b']).toBe(6);
    expect(mapABCD['c']).toBe(3);
    expect(mapABCD['d']).toBe(7);
    expect(pi).toBe(3.141592);
  });
  test('showMessage', () => {
    const spyLog = jest.spyOn(console, 'log');
    const data = 'Hello, World!';
    showMessage(data);
    expect(console.log).toBeCalled();
    expect(spyLog.mock.calls[0][0]).toEqual(data);
    spyLog.mockReset();
    spyLog.mockRestore();
  });
  test('Hamburger', () => {
    const h = new Hamburger();
    jest.spyOn(h, 'listToppings');
    h.listToppings();
    expect(h.listToppings).toBeCalled();
  });
  test('buildName', () => {
    const expected = buildName('First', 'Second', 'Third');
    expect(expected).toBe('First Second Third');
  });
  test('sayName jane', () => {
    const spyLog = jest.spyOn(console, 'log');
    sayName(jane);
    expect(console.log).toBeCalled();
    expect(spyLog.mock.calls[0][0]).toEqual('Hello Jane N/A Doe');
    spyLog.mockReset();
    spyLog.mockRestore();
  });
  test('sayName john', () => {
    const spyLog = jest.spyOn(console, 'log');
    sayName(john);
    expect(console.log).toBeCalled();
    expect(spyLog.mock.calls[0][0]).toEqual('Hello John Smith Doe');
    spyLog.mockReset();
    spyLog.mockRestore();
  });
  test('Test Class', () => {
    const t = new Test();
    expect(t.name).toBe('test');
  });
  test('createDocument', () => {
    // DocumentApp.create = jest.fn().mockImplementation(name => {
    //   return {
    //     getBody: jest.fn().mockImplementation(() => {
    //       return {
    //         appendParagraph: jest.fn().mockImplementation(text => {
    //           this._text = text;
    //         })
    //       };
    //     }),
    //     getId: jest.fn().mockReturnValue("0123456789")
    //   };
    // });
    DocumentApp.create = jest.fn().mockImplementation(name => {
      return new DocumentMock(name);
    });
    const expected = createDocument();
    expect(expected).toBe('0123456789');
  });
});
