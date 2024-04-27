const initialDict = {Red: 2, Yellow: 3, Blue: 5};
const compareDict = {
    Orange: {Red: 6, Yellow: 6},
    Purple: {Red: 5, Blue: 5},
    Green: {Yellow: 2, Blue: 2}
};

function decrementColors(initialDict, compareDict) {
    if (!isColorsCorrect(initialDict, Object.values(compareDict).map(x => Object.keys(x).map(y => y)).flat())) return false

    while (true) {
        if (Object.keys(initialDict).length === 0 ) return true
        const topColor = Object.keys(initialDict).reduce((a, b) => initialDict[a] > initialDict[b] ? a : b);

        if (Object.keys(compareDict).length === 0 ) return false
        const maxKey = getKeyWithMostValue(compareDict, topColor)
        if (!maxKey) return false

        if (initialDict[topColor] >= compareDict[maxKey][topColor]) {
            initialDict[topColor] -= compareDict[maxKey][topColor]
            delete compareDict[maxKey]
            if (initialDict[topColor] === 0) delete initialDict[topColor]
        } else if (initialDict[topColor] < compareDict[maxKey][topColor]) {
            const deductedValue = initialDict[topColor]
            delete initialDict[topColor]
            Object.values(compareDict).forEach(c => c -= deductedValue)
        }
        console.log(initialDict)
    }



}

function getKeyWithMostValue(dict, topColor) {
    let maxKey = null;
    let maxValue = -Infinity;

    for (const key in dict) {
        let totalValue = 0;
        const colorValues = dict[key];
        for (const color in colorValues) {
            totalValue += colorValues[color];
        }
        if (totalValue > maxValue && dict[key].hasOwnProperty(topColor)) {
            maxValue = totalValue;
            maxKey = key;
        }
    }
    return maxKey;
}

function isColorsCorrect(initialDict, compareDict) {
    const initialSet = new Set(Object.keys(initialDict).map(x => x))
    const compareSet = new Set(compareDict)
    for (let item of initialSet) {
        if (!compareSet.has(item)) {
            return false;
        }
    }
    return true;
}

// console.log(isColorsCorrect(initialDict, compareDict))
console.log(decrementColors(initialDict, compareDict));