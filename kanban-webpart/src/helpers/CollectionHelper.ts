export const firstOrDefault = <T>(array: T[], criteria: (T)=>boolean) => {
    let found = array.filter(criteria);
    return found.length && found[0];
};