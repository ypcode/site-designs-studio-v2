export function toJSON(object: any): string {
    return JSON.stringify(object, null, 4);
}