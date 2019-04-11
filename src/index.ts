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

    /**
     * Adds an item to the top of the stack.
     * 
     * @param {T} item the item to be added to the stack
     */
    push(item: T): void {
        this._items.push(item);
    }

    /**
     * Removes an item from the top of the stack, returning it.
     * 
     * @returns {T} the item at the top of the stack.
     */
    pop(): T | any {
        return this._items.pop();
    }

    /**
     * Take a peek at the top of the stack, returning the top-most item,
     * without removing it.
     * 
     * @returns {T} the item at the top of the stack.
     */
    peek(): T | any {
        if (this.isEmpty()) 
            return undefined;
        else
            return this._items[this._items.length - 1];

    }

    /**
     * @returns {boolean} true if the stack is empty.
     */
    isEmpty(): boolean {
        return this._items.length === 0;
    }

    /**
     * @returns {number} the number of items in the stack.
     */
    size(): number {
        return this._items.length;
    }
}
