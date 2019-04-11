/**
 * Implementation of a classic stack.
 */
export class Stack<T> {
    // Internal storage for the stack
    private _items: T[] = [];

    /**
     * Creates a pre-populated stack.
     * 
     * @param {T[]} initialData the initial contents of the stack
     */
    constructor(initialData: Array<T> = []) {
        this._items.push(...initialData)
    }
}
