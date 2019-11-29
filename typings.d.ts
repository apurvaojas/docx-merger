
declare namespace NodeJS {
    interface Global {
        document: Document;
        window: Window | null;
        navigator: Navigator;
    }
}