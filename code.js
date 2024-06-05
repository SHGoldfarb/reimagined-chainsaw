let sheet;

const columns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
const currentValuesStringCoords = [
  "I75",
  "K75",
  "L75",
  "M75",
  "X75",
  "Z75",
  "AA75",
  "AB75",
];

const fromA1Notation = (cell) => {
  const [, columnName, row] = cell.toUpperCase().match(/([A-Z]+)([0-9]+)/);
  const characters = "Z".charCodeAt() - "A".charCodeAt() + 1;

  let column = 0;
  columnName.split("").forEach((char) => {
    column *= characters;
    column += char.charCodeAt() - "A".charCodeAt() + 1;
  });

  return { row: parseInt(row, 10), col: column };
};

const getCell = ({ row, col }) => sheet.getRange(row, col);

const valueOf = (coords) => getCell(coords).getValue();

const estimate = (coords) => {
  const currentValueCoords = fromA1Notation(coords);

  const estimatedValueCoords = {
    ...currentValueCoords,
    row: currentValueCoords.row - 1,
  };
  while (valueOf(estimatedValueCoords) === "") {
    estimatedValueCoords.row -= 1;
  }

  const dailyInterestCoords = {
    ...currentValueCoords,
    row: currentValueCoords.row + 7,
  };

  let changeUnit = 0.0001;
  while (changeUnit > 0.0000000001) {
    while (valueOf(estimatedValueCoords) > valueOf(currentValueCoords)) {
      getCell(dailyInterestCoords).setValue(
        valueOf(dailyInterestCoords) - changeUnit
      );
    }

    while (valueOf(estimatedValueCoords) < valueOf(currentValueCoords)) {
      getCell(dailyInterestCoords).setValue(
        valueOf(dailyInterestCoords) + changeUnit
      );
    }

    changeUnit /= 2;
  }
};

const start = () => {
  // eslint-disable-next-line no-undef
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = spreadsheet.getSheetByName("Rentabilidades");

  currentValuesStringCoords.forEach((coords) => {
    estimate(coords);
  });
};

// eslint-disable-next-line no-unused-vars
function myFunction() {
  start();
}

const sampleArgs = {
  dates: [
    ["2019-08-16T04:00:00.000Z"],
    ["2019-09-02T04:00:00.000Z"],
    ["2019-09-03T04:00:00.000Z"],
    ["2019-09-30T03:00:00.000Z"],
    ["2019-11-04T03:00:00.000Z"],
    ["2019-12-02T03:00:00.000Z"],
    ["2020-01-02T03:00:00.000Z"],
    ["2020-02-03T03:00:00.000Z"],
    ["2020-03-02T03:00:00.000Z"],
    ["2020-03-31T03:00:00.000Z"],
    ["2020-05-01T04:00:00.000Z"],
    ["2020-05-31T04:00:00.000Z"],
    ["2020-06-29T04:00:00.000Z"],
    ["2020-08-04T04:00:00.000Z"],
    ["2020-09-01T04:00:00.000Z"],
    ["2020-09-30T03:00:00.000Z"],
    ["2020-10-31T03:00:00.000Z"],
    ["2020-12-02T03:00:00.000Z"],
    ["2021-01-01T03:00:00.000Z"],
    ["2021-01-31T03:00:00.000Z"],
    ["2021-02-28T03:00:00.000Z"],
    ["2021-03-31T03:00:00.000Z"],
    ["2021-05-01T04:00:00.000Z"],
    ["2021-05-30T04:00:00.000Z"],
    ["2021-06-30T04:00:00.000Z"],
    ["2021-08-01T04:00:00.000Z"],
    ["2021-09-02T04:00:00.000Z"],
    ["2021-09-03T04:00:00.000Z"],
    ["2021-09-30T03:00:00.000Z"],
    ["2021-11-02T03:00:00.000Z"],
    ["2021-11-30T03:00:00.000Z"],
    ["2021-12-30T03:00:00.000Z"],
    ["2022-01-12T03:00:00.000Z"],
    ["2022-01-25T03:00:00.000Z"],
    ["2022-01-28T03:00:00.000Z"],
    ["2022-01-29T03:00:00.000Z"],
    ["2022-02-28T03:00:00.000Z"],
    ["2022-03-21T03:00:00.000Z"],
    ["2022-03-24T03:00:00.000Z"],
    ["2022-03-27T03:00:00.000Z"],
    ["2022-04-01T03:00:00.000Z"],
    ["2022-04-29T04:00:00.000Z"],
    ["2022-05-31T04:00:00.000Z"],
    ["2022-06-29T04:00:00.000Z"],
    ["2022-07-31T04:00:00.000Z"],
    ["2022-08-04T04:00:00.000Z"],
    ["2022-08-30T04:00:00.000Z"],
    ["2022-09-30T03:00:00.000Z"],
    ["2022-11-01T03:00:00.000Z"],
    ["2022-12-02T03:00:00.000Z"],
    ["2022-12-31T03:00:00.000Z"],
    ["2023-01-31T03:00:00.000Z"],
    ["2023-03-02T03:00:00.000Z"],
    ["2023-04-02T04:00:00.000Z"],
    ["2024-01-02T03:00:00.000Z"],
    ["2024-03-18T03:00:00.000Z"],
    ["2024-05-11T04:00:00.000Z"],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
  ],
  movements: [
    [71.51533485446451],
    [7.143701630442742],
    [0],
    [35.65249230530085],
    [7.126225042623734],
    [21.24847542188848],
    [24.72469930350522],
    [31.756113051762465],
    [31.606217434621662],
    [66.43946700161483],
    [20.910593620386994],
    [41.7877932284274],
    [59.2388779877738],
    [41.864739214073246],
    [38.35375903448944],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [-6.810896660243266],
    [-56.88377270021606],
    [-23.38014308681249],
    [-65.84995514062439],
    [0],
    [-19.768394611950384],
    [-19.596977976579158],
    [-62.06641595532212],
    [-29.467753066759496],
    [0],
    [-65.66959301311256],
    [-38.76957886709236],
    [-78.05756053510778],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [0],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
    [""],
  ],
  finalAmount: 0,
};

const milisecondsToDays = (ms) => ms / 1000.0 / 60 / 60 / 24;

const calculateFinalAmount = (dates, movements, dailyReturn) => {
  let currentAmount = 0.0;
  let lastDate = dates[0];
  dates.forEach((date, i) => {
    if (isNaN(date)) {
      return;
    }

    const movement = movements[i] || 0;
    const elapsed = date - lastDate;
    currentAmount *= (1 + dailyReturn) ** elapsed;
    currentAmount += movement;

    lastDate = date;
  });

  return currentAmount;
};

function PROFITABILITY(dates, movements, finalAmount) {
  if (dates.length !== movements.length) {
    throw new Error("dates and movements must be the same size");
  }

  const datesArray = dates.map((dateArr) =>
    milisecondsToDays(Date.parse(dateArr[0]))
  );

  const movementsArray = movements.map((movementArr) => movementArr[0]);

  let step = 0.001;
  let dailyReturn = 0;
  let amount;
  const recalculate = () => {
    amount = calculateFinalAmount(datesArray, movementsArray, dailyReturn);
  };
  recalculate();

  while (Math.abs(amount - finalAmount) > 1) {
    while (amount > finalAmount) {
      dailyReturn -= step;
      recalculate();
    }
    while (amount < finalAmount) {
      dailyReturn += step;
      recalculate();
    }
    step /= 2;
  }

  return dailyReturn;
}
