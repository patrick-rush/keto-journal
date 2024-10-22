const RANGES = {
  carbs: {
    yellowBase: 0,
    greenBase: 10,
    greenCeil: 20,
    yellowCeil: 50,
  },
  fats: {
    yellowBase: 100,
    greenBase: 130,
    greenCeil: 170,
    yellowCeil: 200,
  },
  proteins: {
    yellowBase: 80,
    greenBase: 100,
    greenCeil: 120,
    yellowCeil: 140,
  },
  calories: {
    yellowBase: 1500,
    greenBase: 1800,
    greenCeil: 2000,
    yellowCeil: 2500,
  },
};

const COLS = {
  TIMESTAMP: 1,
  FOOD_ITEM: 2,
  QUANTITY: 3,
  UNITS: 4,
  BRAND_INFO: 5,
  CARBS: 6,
  FATS: 7,
  PROTEINS: 8,
  CALORIES: 9,
  CARBS_TODAY: 10,
  FATS_TODAY: 11,
  PROTEINS_TODAY: 12,
  CALORIES_TODAY: 13,
};

function GET_MACROS(item, quantity, units, brandInfo) {
  if (item === "" || quantity === "") return [[0, 0, 0, 0]];

  const apiKey =
    PropertiesService.getScriptProperties().getProperty("OPEN_AI_API_KEY");
  const url = "https://api.openai.com/v1/chat/completions";

  const headers = {
    Authorization: `Bearer ${apiKey}`,
    "Content-Type": "application/json",
  };

  const queryContent = `Estimate the fat, carbohydrates, fiber, proteins, and calories for ${quantity} ${units} of ${item}`;
  if (brandInfo) queryContent + " (brand information:" + brandInfo + ")";

  const data = {
    model: "gpt-4-0613",
    messages: [
      {
        role: "system",
        content:
          "You are a nutrition expert who provides food macro information in JSON format.",
      },
      {
        role: "user",
        content: `Estimate the fat, carbohydrates, fiber, proteins, and calories for ${quantity} ${units} of ${item}.`,
      },
    ],
    functions: [
      {
        name: "provide_macros",
        description:
          "Provides the macronutrient breakdown for a given food item",
        parameters: {
          type: "object",
          properties: {
            fats: {
              type: "number",
              description: "The amount of fats in grams",
            },
            carbohydrates: {
              type: "number",
              description: "The amount of carbohydrates in grams",
            },
            fiber: {
              type: "number",
              description: "The amount of fiber in grams",
            },
            proteins: {
              type: "number",
              description: "The amount of proteins in grams",
            },
            calories: { type: "number", description: "The amount of calories" },
          },
          required: ["fats", "carbohydrates", "fiber", "proteins", "calories"],
        },
      },
    ],
    function_call: { name: "provide_macros" },
  };

  const options = {
    method: "post",
    headers: headers,
    payload: JSON.stringify(data),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log(response);
    const jsonResponse = JSON.parse(response.getContentText());

    const macroData = jsonResponse.choices[0].message.function_call.arguments;
    const macros = JSON.parse(macroData);
    const carbs = macros.carbohydrates - macros.fiber;

    return [[carbs, macros.fats, macros.proteins, macros.calories]];
  } catch (error) {
    Logger.log("Error fetching macro estimates: " + error);
    return null;
  }
}

function ON_FORM_SUBMIT(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MACROS");

  const row = sheet.getLastRow();

  const foodItem = sheet.getRange(row, COLS.FOOD_ITEM).getValue();
  const quantity = sheet.getRange(row, COLS.QUANTITY).getValue();
  const unit = sheet.getRange(row, COLS.UNITS).getValue() ?? "unit(s)";
  const brandInfo = sheet.getRange(row, COLS.BRAND_INFO).getValue();

  if (foodItem && quantity && unit) {
    const result = GET_MACROS(foodItem, quantity, unit, brandInfo);
    if (result) {
      sheet.getRange(row, COLS.CARBS, 1, 4).setValues(result);
    }

    const today = new Date();
    const dateFormat = "MM/dd/yyyy";

    let totalCarbs = 0,
      totalFats = 0,
      totalProteins = 0,
      totalCalories = 0;

    for (let r = row; r >= 2; r--) {
      const entryDate = sheet.getRange(r, COLS.TIMESTAMP).getValue();
      const formattedEntryDate = Utilities.formatDate(
        entryDate,
        Session.getScriptTimeZone(),
        dateFormat
      );
      const formattedToday = Utilities.formatDate(
        today,
        Session.getScriptTimeZone(),
        dateFormat
      );

      if (formattedEntryDate !== formattedToday) break;

      // Add the macros for the current day (Carbs, Fats, Proteins, Calories)
      totalCarbs += sheet.getRange(r, COLS.CARBS).getValue() ?? 0;
      totalFats += sheet.getRange(r, COLS.FATS).getValue() ?? 0;
      totalProteins += sheet.getRange(r, COLS.PROTEINS).getValue() ?? 0;
      totalCalories += sheet.getRange(r, COLS.CALORIES).getValue() ?? 0;
    }

    const totals = [[totalCarbs, totalFats, totalProteins, totalCalories]];

    sheet.getRange(row, COLS.CARBS_TODAY, 1, 4).setValues(totals);

    if (parseInt(totalCarbs) > RANGES.carbs.greenCeil) {
      const prevRowIndex = row - 1;
      if (prevRowIndex < 1) return;
      const prevRowCarbs = sheet
        .getRange(prevRowIndex, COLS.CARBS_TODAY)
        .getValue();
      if (prevRowCarbs > RANGES.carbs.greenCeil) return;
      SEND_WARNING_NOTIFICATION(
        "Keto Warning!",
        `Warning: You have exceeded the total recommended carb allowance for the day. You have had ${parseFloat(
          totalCarbs
        ).toFixed(2)} carbs today.`
      );
    }
  }
}

function PRODUCE_RECAP() {
  const macrosSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MACROS");
  const recapsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RECAPS");
  const macrosRow = macrosSheet.getLastRow();

  const date = macrosSheet.getRange(macrosRow, COLS.TIMESTAMP).getValue(),
    totalCarbs = macrosSheet.getRange(macrosRow, COLS.CARBS_TODAY).getValue(),
    totalFats = macrosSheet.getRange(macrosRow, COLS.FATS_TODAY).getValue(),
    totalProteins = macrosSheet
      .getRange(macrosRow, COLS.PROTEINS_TODAY)
      .getValue(),
    totalCalories = macrosSheet
      .getRange(macrosRow, COLS.CALORIES_TODAY)
      .getValue();

  const formattedDate = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "MM/dd/yyyy"
  );

  const totals = [
    formattedDate,
    totalCarbs,
    totalFats,
    totalProteins,
    totalCalories,
  ];

  recapsSheet.appendRow(totals);
  SEND_RECAP_NOTIFICATION(...totals);
}

function SEND_RECAP_NOTIFICATION(
  formattedDate,
  totalCarbs,
  totalFats,
  totalProteins,
  totalCalories
) {
  try {
    const email =
      PropertiesService.getScriptProperties().getProperty("EMAIL_ADDRESS");
    GmailApp.sendEmail(
      email,
      `Macros Summary for ${formattedDate}`,
      `
      Macros summary for ${formattedDate}:
        Carbs: ${parseFloat(totalCarbs).toFixed(2)} (${CREATE_EVAL_STRING(
        parseFloat(totalCarbs),
        RANGES.carbs
      )})
        Fats: ${parseFloat(totalFats).toFixed(2)} (${CREATE_EVAL_STRING(
        parseFloat(totalFats),
        RANGES.fats
      )})
        Proteins: ${parseFloat(totalProteins).toFixed(2)} (${CREATE_EVAL_STRING(
        parseFloat(totalProteins),
        RANGES.proteins
      )})
        Calories: ${parseFloat(totalCalories).toFixed(2)} (${CREATE_EVAL_STRING(
        parseFloat(totalCalories),
        RANGES.calories
      )})
      `
    );
  } catch (e) {
    Logger.log("Error sending email notification: " + error);
  }
}

function SEND_WARNING_NOTIFICATION(subject, message) {
  try {
    const email =
      PropertiesService.getScriptProperties().getProperty("EMAIL_ADDRESS");
    GmailApp.sendEmail(email, subject, message);
  } catch (e) {
    Logger.log("Error sending email notification: " + error);
  }
}

function CREATE_EVAL_STRING(total, range) {
  switch (true) {
    case total < range.yellowBase:
      return "Much too low";
    case total >= range.yellowBase && total < range.greenBase:
      return "A little low";
    case total >= range.greenBase && total <= range.greenCeil:
      return "Perfect";
    case total > range.greenCeil && total <= range.yellowCeil:
      return "A little high";
    default:
      return "Much too high";
  }
}

console.log("test change");
