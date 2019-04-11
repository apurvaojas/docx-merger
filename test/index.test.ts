import { expect } from 'chai';
import { Stack } from '../src';

describe('Stack', () => {
    it('can be initialized without an initializer', () => {
        const s = new Stack<number>();
        expect(true).to.be.true;
    });
    
});
