import { useRef, useEffect } from "react";

export const usePrevious = <T extends {}>(value: T): T | undefined => {
    const ref = useRef<T>();
    useEffect(() => {
        ref.current = value;
    });
    return ref.current;
};

export const useCompare = (val: any) => {
    const prevVal = usePrevious(val);
    return prevVal !== val;
};

