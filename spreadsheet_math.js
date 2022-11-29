function loadJS(FILE_URL, then = () => {}, async = true) {
    let scriptEle = document.createElement("script");

    scriptEle.setAttribute("src", FILE_URL);
    scriptEle.setAttribute("type", "text/javascript");
    scriptEle.setAttribute("async", async);

    document.body.appendChild(scriptEle);

    scriptEle.addEventListener("load", () => {
        then();
    });
}

const _FORMULAS_ = [
  ["ACOS" , "acos" ],
  ["ACOSH", "acosh"],
  ["ASIN" , "asin" ],
  ["ASINH", "asinh"],
  ["ATAN" , "atan" ],
  ["ATANH", "atan" ],
  ["COS"  , "cos"  ],
  ["COSH" , "cosh" ],
  ["CSC"  , "csc"  ],
  ["PI()" , "pi"   ],
  ["SEC"  , "sec"  ],
  ["SIN"  , "sin"  ],
  ["SQRT" , "sqrt" ],
  ["TAN"  , "tan"  ],
];

function preparse(value) {
    value = value.toString();
    for (let formula of _FORMULAS_) {
        value = value.replace(formula[0], formula[1]);
    }
    return value;
}

function postparse(value) {
    for (let formula of _FORMULAS_) {
        value = value.replace(formula[1], formula[0]);
    }
    return value;
}


function linearRegression(x, y) {
    const sumX = x.reduce((prev, curr) => prev + curr, 0);
    const avgX = sumX / x.length;
    const xDifferencesToAverage = x.map((value) => avgX - value);
    const xDifferencesToAverageSquared = xDifferencesToAverage.map(
      (value) => value ** 2
    );
    const SSxx = xDifferencesToAverageSquared.reduce(
      (prev, curr) => prev + curr,
      0
    );
    const sumY = y.reduce((prev, curr) => prev + curr, 0);
    const avgY = sumY / y.length;
    const yDifferencesToAverage = y.map((value) => avgY - value);
    const xAndYDifferencesMultiplied = xDifferencesToAverage.map(
      (curr, index) => curr * yDifferencesToAverage[index]
    );
    const SSxy = xAndYDifferencesMultiplied.reduce(
      (prev, curr) => prev + curr,
      0
    );
    const slope = SSxy / SSxx;
    const intercept = avgY - slope * avgX;
    return [slope, intercept];
}

function factorial(n) {
    return n > 1 ? n * factorial(n-1) : 1.;
}



( function () { 'use strict';

    const { args, toString, toNumber } = o_spreadsheet.helpers;

    loadJS("https://unpkg.com/mathjs@11.4.0/lib/browser/math.js", () => {
        o_spreadsheet.addFunction("EVALUATE", {
            description: "Evaluation of an expression with variable",
            args: args(`
                expression (string) "The first number or range to add together."
                variable (string,optional,repeating) "Field name."
                value (number,optional,repeating) "Value."
            `),
            returns: ["NUMBER"],
            compute: function (value, ...values) {
                value = preparse(value);
                let scope = {};
                for (let i = 0; i < values.length; i += 2) {
                    scope[toString(values[i])] = toNumber(values[i + 1]);
                }
                return math.evaluate(value, scope);
            },
        });

        o_spreadsheet.addFunction("SIMPLIFY", {
            description: "Simplify an expression",
            args: args(`
                expression (string) "The first number or range to add together."
                variable (string,optional,repeating) "Field name."
                value (number,optional,repeating) "Value."
            `),
            returns: ["STRING"],
            compute: function (value, ...values) {
                value = preparse(value);
                let scope = {};
                for (let i = 0; i < values.length; i += 2) {
                    scope[toString(values[i])] = toNumber(values[i + 1]);
                }
                let computed = math.simplify(value, scope).toString();
                return postparse(computed);
            },
        });

        o_spreadsheet.addFunction("DERIVATIVE", {
            description: "Differentiate an expression",
            args: args(`
                expression (string) "The first number or range to add together."
                variable (string, optional) "Field name."
                value (number,optional) "Value."
            `),
            returns: ["STRING"],
            compute: function (expression, variable = "x", value) {
                expression = preparse(expression);
                const diff = math.derivative(expression, variable);
                return value ? diff.evaluate({[variable]:value}) : postparse(diff.toString());
            },
        });

        o_spreadsheet.addFunction("TAYLOR", {
            description: "Compute the definite integrate of a function",
            args: args(`
                expression (string) "The function to compute the limit"
                variable (string) "The variable."
                order (number) "The limit value."
                point (number, default=0.) "The starting point"
            `),
            returns: ["NUMBER"],
            compute: function (expression, variable, order, point = 0.) {
                expression = preparse(expression);
                const context = {[variable] : point};
                let solution = `${math.evaluate(expression, context)}`;
                for (let i = 1; i < order+1; ++i) {
                    expression = math.derivative(expression, variable);
                    solution += ` + (${expression.evaluate(context)}/${factorial(i)}) * (${variable} - ${point}) ^ ${i}`;
                }
                return postparse(math.simplify(solution).toString());
            },
        });
    });

    loadJS("https://cdn.jsdelivr.net/npm/nerdamer@1.1.13/all.min.js", () => {
        o_spreadsheet.addFunction("INTEGRATE", {
            description: "Differentiate an expression",
            args: args(`
                expression (string) "The first number or range to add together."
                variable (string, optional) "Field name."
            `),
            returns: ["STRING"],
            compute: function (value, variable = "x") {
                value = preparse(value);
                let computed = nerdamer.integrate(value, variable).toString();
                return postparse(computed);
            },
        });

        o_spreadsheet.addFunction("DEFINT", {
            description: "Compute the definite integrate of a function",
            args: args(`
                expression (string) "To function to be integrated."
                variable (string) "The variable to integrate on."
                start (number) "The lower bound of the integration domain."
                stop (number) "The uppoer bound of the integration domain."
            `),
            returns: ["NUMBER"],
            compute: function (value, variable, start, stop) {
                value = preparse(value);
                return nerdamer.defint(value, start, stop, variable).text();
            },
        });

        o_spreadsheet.addFunction("LIMIT", {
            description: "Compute the definite integrate of a function",
            args: args(`
                expression (string) "The function to compute the limit"
                variable (string) "The variable."
                value (number) "The limit value."
            `),
            returns: ["NUMBER"],
            compute: function (value, variable, point) {
                value = preparse(value);
                return nerdamer.limit(value, variable, point).toString();
            },
        });

        o_spreadsheet.addFunction("EXPAND", {
            description: "Expand an expression",
            args: args(`
                expression (string) "The function to compute the limit"
            `),
            returns: ["NUMBER"],
            compute: function (expression) {
                expression = preparse(expression);
                return postparse(nerdamer.expand(expression).toString());
            },
        });
    });

    // CORRELATION COEFFICIENTS =============================================================================================

    o_spreadsheet.addFunction("PEARSON", {
        description: "Compute the Pearson Coefficient",
        args: args(`
            value1 (number, range<number>) "The first value or range to consider when calculating the average value."
            value2 (number, range<number>) "Additional values or ranges to consider when calculating the average value."
        `),
        returns: ["NUMBER"],
        compute: function (d1, d2) {
            let { min, pow, sqrt } = Math
            let add = (a, b) => a + b
            const n = min(d1[0].length, d2[0].length)
            if (n === 0) {
                return 0
            }
            [d1, d2] = [d1[0].slice(0, n), d2[0].slice(0, n)]
            let [sum1, sum2] = [d1, d2].map(l => l.reduce(add))
            let [pow1, pow2] = [d1, d2].map(l => l.reduce((a, b) => a + pow(b, 2), 0))
            let mulSum = d1.map((n, i) => n * d2[i]).reduce(add)
            let dense = sqrt((pow1 - pow(sum1, 2) / n) * (pow2 - pow(sum2, 2) / n))
            if (dense === 0) {
                return 0
            }
            return (mulSum - (sum1 * sum2 / n)) / dense
        },
    });

    o_spreadsheet.addFunction("SPEARMAN", {
        description: "Compute the Spearman Coefficient",
        args: args(`
            value1 (number, range<number>) "The first value or range to consider when calculating the average value."
            value2 (number, range<number>) "Additional values or ranges to consider when calculating the average value."
        `),
        returns: ["NUMBER"],
        compute: function (d1, d2) {
            const n = Math.min(d1[0].length, d2[0].length);
            if (n === 0) {
                return 0
            }
            [d1, d2] = [d1[0].slice(0, n), d2[0].slice(0, n)]
            let order = d1.map((e,i) => [e, d2[i]]);
            order.sort((a,b) => a[0] - b[0]);

            for(let i = 0; i < n; ++i){
                order[i].push(i+1);
            }
            order.sort((a,b) => a[1] - b[1]);

            for(let i = 0; i < n; ++i){
                order[i].push(i+1);
            }
            const sum = order
                .map(o => (o[2]-o[3])**2.)
                .reduce((a, b) => a + b);

            return 1 - (6 * sum / (n**3. - n));
        },
    });

    o_spreadsheet.addFunction("MSE", {
        description: "Compute the mean squared error",
        args: args(`
            value1 (number, range<number>) "The first value or range to consider when calculating the average value."
            value2 (number, range<number>) "Additional values or ranges to consider when calculating the average value."
        `),
        returns: ["NUMBER"],
        compute: function (values1, values2) {
            let error = 0.0;
            for (let i = 0; i < values1[0].length; ++i) {
                error += (values1[0][i]-values2[0][i])**2.;
            }
            return error/values1[0].length;
        },
    });

    o_spreadsheet.addFunction("MATTHEWS", {
        description: "Compute the matthews coefficient score",
        args: args(`
            value1 (number, range<number>) "The first value or range to consider when calculating the average value."
            value2 (number, range<number>) "Additional values or ranges to consider when calculating the average value."
        `),
        returns: ["NUMBER"],
        compute: function (values1, values2) {
            let TN = TP = FP = FN = 0;
            for (let i = 0; i < values1[0].length; ++i) {
                const b1 = !!values1[0][i];
                const b2 = !!values2[0][i];
                TP += (!b1 && !b2);
                TN += (b1 && b2);
                FN += (!b1 && b2);
                FP += (b1 && !b2);
            }
            return  (TP * TN - FP * FN) / Math.sqrt((TP + FP) * (TP + FN) * (TN + FP) * (TN + FN));
        },
    });

    // MODELISATION ============================================================================================
    o_spreadsheet.addFunction("LINREG", {
        description: "Compute the Linear Regression",
        args: args(`
            value1 (number, range<number>) "The first value or range to consider when calculating the average value."
            value2 (number, range<number>) "Additional values or ranges to consider when calculating the average value."
        `),
        returns: ["STRING"],
        compute: function (d1, d2) {
            const [a, b] = linearRegression(d1[0], d2[0]);
            return `${a}*X + ${b}`;
        },
    });
    o_spreadsheet.addFunction("LINREG.SLOPE", {
        description: "Compute the Linear Regression",
        args: args(`
            value1 (number, range<number>) "The first value or range to consider when calculating the average value."
            value2 (number, range<number>) "Additional values or ranges to consider when calculating the average value."
        `),
        returns: ["NUMBER"],
        compute: function (d1, d2) {
            const [a, b] = linearRegression(d1[0], d2[0]);
            return a;
        },
    });
    o_spreadsheet.addFunction("LINREG.INTERCEPT", {
        description: "Compute the Linear Regression",
        args: args(`
            value1 (number, range<number>) "The first value or range to consider when calculating the average value."
            value2 (number, range<number>) "Additional values or ranges to consider when calculating the average value."
        `),
        returns: ["NUMBER"],
        compute: function (d1, d2) {
            const [a, b] = linearRegression(d1[0], d2[0]);
            return b;
        },
    });

    loadJS("https://cdnjs.cloudflare.com/ajax/libs/numeric/1.2.6/numeric.js", () => {
        o_spreadsheet.addFunction("POLYFIT", {
            description: "Differentiate an expression",
            args: args(`
                values1 (range<number>) "The first number or range to add together."
                values2 (range<number>) "Field name."
                order (number) "The order"
            `),
            returns: ["STRING"],
            compute: function (x, y, order) {
                x = x[0];
                y = y[0];
                var xMatrix = [];
                var xTemp = [];
                var yMatrix = numeric.transpose([y]);

                for (let j=0;j<x.length;j++)
                {
                    xTemp = [];
                    for(let i=0;i<=order;i++)
                    {
                        xTemp.push(1*Math.pow(x[j],i));
                    }
                    xMatrix.push(xTemp);
                }

                var xMatrixT = numeric.transpose(xMatrix);
                var dot1 = numeric.dot(xMatrixT,xMatrix);
                var dotInv = numeric.inv(dot1);
                var dot2 = numeric.dot(xMatrixT,yMatrix);
                var solution = numeric.dot(dotInv,dot2);

                let poly = `${solution[0]}`;
                for (let i = 1; i < solution.length; ++i) {
                    poly += ` + (${solution[i]}*X^${i})`;
                }
                return poly;
            },
        });
    });

    o_spreadsheet.addFunction("LOG", {
        description: "The logarithm of a number, base 10.",
        args: args(`
        value (number) "The value for which to calculate the logarithm, base e."
        basis (number, default=10.) "The value for which to calculate the logarithm, base e."
        `),
        returns: ["NUMBER"],
        compute: function (value, basis = 10) {
            const _value = toNumber(value);
            assert(() => _value > 0, _lt("The value (%s) must be strictly positive.", _value.toString()));
            return Math.log(_value) / Math.log(toNumber(basis));
        },
    });

})();
