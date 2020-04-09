export const getTrimmedText = (text: string, trimTo: number) => {
    if (text && text.length > trimTo) {
        return `${text.substring(0, trimTo)}...`;
    }

    return text;
};
