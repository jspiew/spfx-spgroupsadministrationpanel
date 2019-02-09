export type Draft<T> = {
    [P in keyof T]?: T[P];
};