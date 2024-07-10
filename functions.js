// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Adds two numbers without using batching
 * @CustomFunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function addNoBatch(first, second) {
  return first + second;
}

/**
 * Sums total tax payable for a given year
 * @CustomFunction
 * @param first year
 * @param taxableIncome taxable income
 * @param SAPTO seniors and pensioners tax offset
 * @param familyIncome family income
 * @param dependents number of dependents
 * @returns Sum of income tax, medicare levy, LITO, and LAMITO
 */
async function totalTax(year, taxableIncome, SAPTO = 0, familyIncome = -1, dependents = 0) {
  try {
    // Run all tax calculations in parallel and await their results
    const results = await Promise.all([
        incomeTax(year, taxableIncome),
        medicareLevy(year, taxableIncome, SAPTO, familyIncome, dependents),
        lito(year, taxableIncome),
        lamito(year, taxableIncome)
    ]);

    // Check for any errors (string results)
    const errors = results.filter(result => typeof result === 'string');
    if (errors.length > 0) {
        // Concatenate all string errors and return
        return 'Errors: ' + errors.join("; ");
    }

    // Calculate the total tax using a loop
    let totalTax = 0;
    for (const result of results) {
        if (typeof result === 'number') {
            totalTax += result;
        } else {
            // If we encounter any string after the check (unlikely), handle it
            console.error('Unexpected string in results:', result);
            return 'Unexpected error occurred'; // or handle this case as appropriate
        }
    }

    // Ensure the total tax is not negative
    return Math.max(0, totalTax);
} catch (ex) {
    console.error('Error calculating total tax:', ex);
    return 'Error calculating total tax';  // Return an error message in case of an exception
}
}


/**
 * Returns the income tax
 * @CustomFunction
 * @param year year
 * @param taxableIncome taxable income
 * @returns income tax
 */
async function incomeTax(year, taxableIncome) {
    try {
        console.log('Loading tax rates...');
        const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/tax_rates.json");
        const rates = await response.json();
        console.log('Rates loaded:', rates);

        let rateKey = Object.keys(rates).find(key => {
            if (key.includes('_')) {
                const [start, end] = key.split('_').map(Number);
                return year >= start && year <= end;
            }
            return key === year.toString();
        });

        console.log('Rate key found:', rateKey);

        if (!rateKey) {
            return "#YearError";
        }

        const taxBrackets = rates[rateKey].brackets;
        console.log('Tax brackets:', taxBrackets);
        let retText = 0;

        for (let i = taxBrackets.length - 1; i >= 0; i--) {
            if (taxableIncome > taxBrackets[i].threshold) {
                retText = taxBrackets[i].baseTax + taxBrackets[i].rate * (taxableIncome - taxBrackets[i].threshold);
                break;
            }
        }

        return retText;
    } catch (error) {
        console.log("test");
        console.error('Failed to calculate tax:', error);
        return "Erroraa"; // Error handling
    }
}

/**
 * Returns the income minor tax
 * @CustomFunction
 * @param year year
 * @param eligibleIncome eligible income
 * @returns minor tax
 */
async function incomeTaxMinor(year, eligibleIncome) {
  try {
    // Load tax rates from JSON file
    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/income_tax_minor.json");
    const data = await response.json();
    const taxRates = data.rates;

    let retText = 0; // Declare retText to handle both number and string

    // Check if the year is within the range 2015-2030
    if (year >= 2015 && year <= 2030) {
      // Calculate tax based on rates
      for (let i = taxRates.length - 1; i >= 0; i--) {
        if (eligibleIncome > taxRates[i].threshold) {
          retText = taxRates[i].baseTax + taxRates[i].rate * (eligibleIncome - taxRates[i].threshold);
          break; // Exit loop once the correct bracket is found
        }
      }
    } else {
      retText = "#YearError"; // Use string for error message
    }

    return retText; // Return either the calculated tax or an error message
  } catch (ex) {
    console.error("Error fetching or calculating tax:", ex);
    return "Error"; // Return error as a string in case of an exception
  }
}

/**
 * Returns the medicare levy
 * @CustomFunction
 * @param year year
 * @param taxableIncome taxable income
 * @param SAPTO seniors and pensioners tax offset
 * @param familyIncome family income
 * @param dependents number of dependents
 * @returns medicare levy
 */
export async function medicareLevy(year, taxableIncome, SAPTO = 0, familyIncome = -1, dependents = 0) {
  try {
    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/medicare_rates.json");
    const rates = await response.json();

    // Find the latest year available in the data
    const years = Object.keys(rates).map(Number);
    const latestYear = Math.max(...years);

    // Check if the year is not present in the data and use the latest available year
    if (year > latestYear) {
      year = latestYear;
    } else if (!rates[year]) {
      return "#YearError"; // Return an error if the year is not valid
    }

    const data = rates[year];
    const rate = data.Rate;
    const stepin = data.Stepin;
    let lower, upper, familyLower, familyUpper;

    if (SAPTO == -1) {
      lower = data.SAPTO.Lower;
      upper = data.SAPTO.Upper;
      familyLower = data.SAPTO.FamilyLowerBase + data.SAPTO.FamilyStepLower * dependents;
      familyUpper = data.SAPTO.FamilyUpperBase + data.SAPTO.FamilyStepUpper * dependents;
    } else {
      lower = data.Lower;
      upper = data.Upper;
      familyLower = data.FamilyLowerBase + data.FamilyStepLower * dependents;
      familyUpper = data.FamilyUpperBase + data.FamilyStepUpper * dependents;
    }

    let retText = calculateTax(taxableIncome, familyIncome, rate, stepin, lower, upper, familyLower, familyUpper);
    return retText;
  } catch (ex) {
    return null; // Handle errors or exceptions
  }
}

function calculateTax(taxableIncome, familyIncome, rate, stepin, lower, upper, familyLower, familyUpper) {
  let retText = 0;
  let SY, reduction, sreduction;

  if (familyIncome > 0) {
    // Taxpayer has a spouse or dependent children
    SY = Math.max(familyIncome - taxableIncome, 0);

    if (SY === 0 || taxableIncome === 0) {
      if (familyIncome > familyUpper) {
        retText = Math.max(taxableIncome, SY) * rate;
      } else if (familyIncome > familyLower) {
        retText = (Math.max(taxableIncome, SY) - familyLower) * stepin;
      }
    } else if (familyIncome > familyUpper) {
      if (taxableIncome > upper) {
        retText = taxableIncome * rate;
      } else if (taxableIncome > lower) {
        retText = (taxableIncome - lower) * stepin;
      }
    } else if (SY > upper && taxableIncome > upper) {
      retText = (taxableIncome / familyIncome) * stepin * (familyIncome - familyLower);
    } else if (SY < lower && taxableIncome > lower) {
      if (taxableIncome > upper) {
        retText = taxableIncome * rate;
      } else if (taxableIncome > lower) {
        retText = (taxableIncome - lower) * stepin;
      }
      retText = Math.max(retText - (rate * familyLower - 0.08 * (familyIncome - familyLower)), 0);
    } else if (SY > upper && taxableIncome > lower && taxableIncome < upper) {
      if (taxableIncome > upper) {
        retText = taxableIncome * rate;
      } else if (taxableIncome > lower) {
        retText = (taxableIncome - lower) * stepin;
      }
      if (SY > upper) {
        sreduction = taxableIncome * rate;
      } else if (SY > lower) {
        sreduction = (taxableIncome - lower) * stepin;
      }
      reduction = rate * familyLower - 0.08 * (familyIncome - familyLower);
      retText = retText - (reduction * taxableIncome) / familyIncome;
      reduction = Math.min((reduction * SY) / familyIncome, 0);
      retText = Math.max(retText, 0) + reduction;
    } else if (SY > lower && SY < upper && taxableIncome > upper) {
      sreduction = 0;
      if (taxableIncome > upper) {
        retText = taxableIncome * rate;
      } else if (taxableIncome > lower) {
        retText = (taxableIncome - lower) * stepin;
      }
      if (SY > upper) {
        sreduction = SY * rate;
      } else if (SY > lower) {
        sreduction = (SY - lower) * stepin;
      }
      reduction = rate * familyLower - 0.08 * (familyIncome - familyLower);
      retText = retText - (reduction * taxableIncome) / familyIncome;
      sreduction = Math.min(sreduction - (reduction * SY) / familyIncome, 0);
      retText = Math.max(retText, 0) + sreduction;
    }
  } else {
    if (taxableIncome > upper) {
      retText = taxableIncome * rate;
    } else if (taxableIncome > lower) {
      retText = (taxableIncome - lower) * stepin;
    }
  }

  return retText;
}

/**
 * Returns the lito
 * @customfunction
 * @returns lito
 */
export async function lito(year, taxableIncome) {
  try {
    if (taxableIncome <= 0) {
      return 0;
    }

    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/lito_rates.json");
    const rates = await response.json();

    let configKey = Object.keys(rates).find((key) => {
      if (key.includes("_")) {
        const [start, end] = key.split("_").map(Number);
        return year >= start && year <= end;
      }
      return key === year.toString();
    });

    console.log("Rate key found:", configKey);

    if (!configKey) {
      return "#YearError";
    }

    const config = rates[configKey];
    let retText = 0;

    if (config.offset !== undefined && config.Stepin !== undefined && config.Lower !== undefined) {
      // Handling the years 2013 to 2020
      if (taxableIncome <= config.Lower) {
        retText = -config.offset;
      } else {
        retText = -Math.max(config.offset - (taxableIncome - config.Lower) * config.Stepin, 0);
      }
    } else if (config.brackets) {
      // Handling the years 2021 to 2030
      for (const bracket of config.brackets) {
        if (taxableIncome <= bracket.threshold) {
          retText = -bracket.offset;
          break;
        } else {
          retText = -Math.max(bracket.offset - (taxableIncome - bracket.threshold) * bracket.rate, 0);
        }
      }
    }

    return retText;
  } catch (error) {
    console.error("Error fetching or calculating LITO:", error);
    return "Error"; // Error handling
  }
}


/**
 * Returns the mito
 * @customfunction
 * @returns mito
 */
export async function lamito(year, taxableIncome) {
  try {
    if (taxableIncome <= 0) {
      return 0;
    }

    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/mito_rates.json");
    const rates = await response.json();

    let yearKey = Object.keys(rates).find((key) => {
      if (key.includes("_")) {
        const [start, end] = key.split("_").map(Number);
        return year >= start && year <= end;
      }
      return key === year.toString();
    });

    console.log("Rate key found:", yearKey);

    if (!yearKey) {
      return "#YearError";
    }

    const brackets = rates[yearKey].brackets;
    let retText = 0;

    // Iterate through brackets from the highest to the lowest
    for (let i = brackets.length - 1; i >= 0; i--) {
      if (taxableIncome > brackets[i].threshold) {
        const increment = (taxableIncome - brackets[i].threshold) * brackets[i].incrementRate;
        const decrement = (taxableIncome - brackets[i].threshold) * brackets[i].decrementRate;
        retText = -Math.max(brackets[i].offset + increment - decrement, 0);
        break;
      }
    }

    return retText;
  } catch (error) {
    console.error("Error fetching or calculating LAMITO:", error);
    return "Error"; // Error handling
  }
}

/**
 * Returns help debt
 * @customfunction
 * @returns help
 */
export async function help(year, taxableIncome, helpDebt) {
  try {
    // Load the HELP repayment rates
    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/help_rates.json");
    const rates = await response.json();

    // Determine the appropriate rates for the given year
    const rateKey = Object.keys(rates).find((key) => key.includes(year.toString()));
    if (!rateKey) {
      return "#YearError";
    }

    const yearlyRates = rates[rateKey];
    let retText = 0;

    // Determine the repayment amount
    for (const rate of yearlyRates) {
      if (taxableIncome > rate.threshold) {
        retText = Math.min(taxableIncome * rate.rate, helpDebt);
      } else {
        break; // No need to continue if the income is below the current threshold
      }
    }

    return retText;
  } catch (error) {
    console.error("Failed to calculate HELP debt:", error);
    return "Error"; // Error handling
  }
}


/**
 * Returns ftba
 * @customfunction
 * @returns ftba
 */
export async function ftba(
  year,
  familyAdjustedIncome,
  age,
  options,
  maxRate = 0,
  baseRate = 0
) {
  try {
    if (age === -1) return 0;

    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/ftbA_rates.json");
    const config = await response.json();

    const yearData = config.years[year.toString()];
    if (!yearData) return "#YearError";

    const { LowerLimit, UpperLimit, maxRates, baseRates, surcharges } = yearData;

    let retText = 0;

    switch (options) {
      case 1: // Maximum rate payable (including EOY supplement if relevant)
        retText = age <= 12 ? maxRates["0-12"] : maxRates["13-19"];
        if (familyAdjustedIncome > 80000) {
          retText -= surcharges.maxRate;
        }
        break;
      case 2: // Base rate payable (including EOY supplement if relevant)
        retText = baseRates["0-19"];
        if (familyAdjustedIncome > 80000) {
          retText -= surcharges.baseRate;
        }
        break;
      case 3: // Lower income test threshold
        retText = LowerLimit;
        break;
      case 4: // Upper income test threshold
        retText = UpperLimit;
        break;
      case 5: // Income Test 1
        retText =
          maxRate -
          Math.min(
            Math.max(Math.min(familyAdjustedIncome, UpperLimit) - LowerLimit, 0) * 0.2 +
              Math.max(familyAdjustedIncome - UpperLimit, 0) * 0.3,
            maxRate
          );
        break;
      case 6: // Income Test 2
        retText = baseRate - Math.min(Math.max(familyAdjustedIncome - UpperLimit, 0) * 0.3, baseRate);
        break;
      case 7: // FTB Part A Amount
        retText = Math.max(
          baseRate - Math.min(Math.max(familyAdjustedIncome - UpperLimit, 0) * 0.3, baseRate),
          maxRate -
            Math.min(
              Math.max(Math.min(familyAdjustedIncome, UpperLimit) - LowerLimit, 0) * 0.2 +
                Math.max(familyAdjustedIncome - UpperLimit, 0) * 0.3,
              maxRate
            )
        );
        break;
      default:
        return "#OptionError";
    }

    return retText;
  } catch (ex) {
    console.error("Failed to calculate SDFTBA:", ex);
    return "Error";
  }
}

/**
 * Returns ftbb2
 * @customfunction
 * @returns ftbb2
 */
export async function ftbb(
  year,
  familyAdjustedIncome,
  minimumAdjustedIncome,
  age,
  options,
  status
) {
  try {
    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/ftbB_rates.json");
    const config = await response.json();

    const yearData = config[year.toString()];
    if (!yearData) return "#YearError";

    const { ageRates, incomeThreshold, coupleThreshold } = yearData;
    let maxAmt;
    let retText = 0;

    if (age === null || age < 0) return 0;

    if (age <= 4) {
      maxAmt = ageRates["0-4"];
    } else if (age <= 18) {
      maxAmt = ageRates["5-18"];
    } else {
      return 0; // If age doesn't fall into valid range
    }

    switch (options) {
      case 1: // Maximum rate payable (including EOY supplement if relevant)
        retText = maxAmt;
        break;
      case 2: // Lower Income Threshold
        retText = incomeThreshold;
        break;
      case 3: // FTB Part B amount
        if (status === "Couple") {
          if (age < 14 && familyAdjustedIncome - minimumAdjustedIncome <= coupleThreshold) {
            retText = maxAmt - Math.min(Math.max(minimumAdjustedIncome - incomeThreshold, 0) * 0.2, maxAmt);
          } else {
            retText = 0;
          }
        } else if (status === "Single") {
          if (age < 19 && minimumAdjustedIncome <= coupleThreshold) {
            retText = maxAmt;
          } else {
            retText = 0;
          }
        } else {
          retText = 0;
        }
        break;
      default:
        return "#OptionError";
    }

    return retText;
  } catch (ex) {
    console.error("Failed to calculate SDFTBB:", ex);
    return "Error";
  }
}

/**
 * Returns div293
 * @customfunction
 * @returns div293
 */
export async function div2932(year, income, superAmount) {
  try {
    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/div293.json");
    const config = await response.json();

    const threshold = config.thresholds[year.toString()];
    if (threshold === undefined) return "#YearError";

    let retText = 0;

    if (income > threshold) {
      retText = Math.min(superAmount * 0.15, (income - threshold) * 0.15);
    } else {
      retText = 0;
    }

    return retText;
  } catch (ex) {
    console.error("Failed to calculate SDDiv293:", ex);
    return "Error";
  }
}

/**
 * Returns healthins
 * @customfunction
 * @returns healthins
 */
export async function SDHealthInsTier(
  year,
  surchargeIncome,
  age = 0,
  familyIncome = -1,
  dependents = 0
) {
  try {
    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/health_ins_tiers.json");
    const config = await response.json();

    let yearData;
    for (const range in config.ranges) {
      const [startYear, endYear] = range.split("_").map(Number);
      if (year >= startYear && year <= endYear) {
        yearData = config.ranges[range];
        break;
      }
    }

    if (!yearData) return "#YearError";

    let income;
    let rebates;

    if (familyIncome === -1) {
      // Taxpayer is single
      rebates = yearData.single;
      income = surchargeIncome;
    } else {
      // Taxpayer has a spouse or dependent children
      rebates = yearData.family.map((r) => ({
        ...r,
        threshold: (r.baseThreshold ?? 0) + Math.max(dependents - 1, 0) * 1500,
      }));
      income = familyIncome;
    }

    let retText = "#YearError";

    for (let i = 0; i < rebates.length; i++) {
      if (income >= (rebates[i].threshold ?? 0)) {
        retText = rebates[i].tier;
      }
    }

    return retText;
  } catch (ex) {
    console.error("Failed to calculate SDHealthInsTier:", ex);
    return "Error";
  }
}


/**
 * Returns sgc
 * @customfunction
 * @returns sgc
 */
export async function sgc(year) {
  try {
    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/SGC.json");
    const config = await response.json();

    const rate = config.rates[year.toString()];
    if (rate === undefined) return "#YearError";

    return rate;
  } catch (ex) {
    console.error("Failed to calculate SDSGC:", ex);
    return "Error";
  }
}


/**
 * Returns medicarelevysurcharge
 * @customfunction
 * @returns medicarelevysurcharge
 */
export async function medicareLevySurcharge(
  year,
  surchargeIncome,
  sapto = 0,
  familyIncome = -1,
  dependents = 0
) {
  try {
    const response = await fetch("https://sd360.z26.web.core.windows.net/datafiles/medicare_levy_surcharge.json");
    const config = await response.json();

    let yearData;
    for (const range in config.ranges) {
      const [startYear, endYear] = range.split("_").map(Number);
      if (year >= startYear && year <= endYear) {
        yearData = config.ranges[range];
        break;
      }
    }

    if (!yearData) return "#YearError";

    let income;
    let surcharges;

    if (familyIncome === -1) {
      // Taxpayer is single
      surcharges = yearData.single;
      income = surchargeIncome;
    } else {
      // Taxpayer has a spouse or dependent children
      surcharges = yearData.family.map((s) => ({
        ...s,
        threshold: (s.baseThreshold ?? 0) + Math.max(dependents - 1, 0) * 1500,
      }));
      income = surchargeIncome;
    }

    if (income <= 22398) {
      return 0;
    }

    let retText = 0;

    for (let i = 0; i < surcharges.length; i++) {
      if (income >= (surcharges[i].threshold ?? 0)) {
        retText = income * surcharges[i].rate;
      }
    }

    return retText;
  } catch (ex) {
    console.error("Failed to calculate SDMedicareLevySurcharge:", ex);
    return "Error";
  }
}


/**
 * Defines the implementation of the custom functions
 * for the function id defined in the metadata file (functions.json).
 */
CustomFunctions.associate("ADDNOBATCH", addNoBatch);
CustomFunctions.associate("TotalTax", totalTax);
CustomFunctions.associate("IncomeTax", incomeTax);
CustomFunctions.associate("IncomeTaxMinor", incomeTaxMinor);
CustomFunctions.associate("MedicareLevy", medicareLevy);
CustomFunctions.associate("LITO", lito);
CustomFunctions.associate("LAMITO", lamito);
CustomFunctions.associate("help", help);
CustomFunctions.associate("ftba", ftba);
CustomFunctions.associate("ftbb", ftbb);
CustomFunctions.associate("DIV2932", div2932);
CustomFunctions.associate("SDHealthInsTier", SDHealthInsTier);
CustomFunctions.associate("SGC", sgc);
CustomFunctions.associate("MedicareLevySurcharge", medicareLevySurcharge);





///////////////////////////////////////

let _batch = [];
let _isBatchedRequestScheduled = false;

// This function encloses your custom functions as individual entries,
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.
function _pushOperation(op, args) {
  // Create an entry for your custom function.
  console.log("pushOperation");
  const invocationEntry = {
    operation: op, // e.g. sum
    args: args,
    resolve: undefined,
    reject: undefined,
  };

  // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g. 100 ms.
  if (!_isBatchedRequestScheduled) {
    console.log("schedule remote request");
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}

// This is a private helper function, used only within your custom function add-in.
// You wouldn't call _makeRemoteRequest in Excel, for example.
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  try{
  console.log("makeRemoteRequest");
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });
  console.log("makeRemoteRequest2");
  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      console.log("responseBatch in fetchFromRemoteService");
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
          console.log("rejecting promise");
        } else {
          console.log("fulfilling promise");
          console.log(response);

          batchCopy[index].resolve(response.result);
        }
      });
    });
    console.log("makeRemoteRequest3");
  } catch (error) {
    console.log("error name:" + error.name);
    console.log("error message:" + error.message);
    console.log(error);
  }
}

// --------------------- A public API ------------------------------

// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a [promise of] batch of values.
// NOTE: When implementing this function on a server, also apply an appropriate authentication mechanism
//       to ensure only the correct callers can access it.
async function _fetchFromRemoteService(requestBatch) {
  // Simulate a slow network request to the server;
  console.log("_fetchFromRemoteService");
  await pause(1000);
  console.log("postpause");
  return requestBatch.map((request) => {
    console.log("requestBatch server side");
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myResult = args[0] * args[1];
        console.log(myResult);
        return {
          result: myResult
        };
      } else {
        return {
          error: `Operation not supported: ${operation}`
        };
      }
    } catch (error) {
      return {
        error: `Operation failed: ${operation}`
      };
    }
  });
}

function pause(ms) {
  console.log("pause");
  return new Promise((resolve) => setTimeout(resolve, ms));
}
