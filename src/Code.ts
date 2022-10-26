function runScript() {
  const sheetName = {
    input: 'script-input',
    outpuSolveTracker: 'script-output: Solve Tracker'
  } as const;
  const tableHeadingInSheet = {
    contest: 'input-contest',
    user: 'input-user',
    color: 'input-color'
  } as const;
  const baseUrlVjudgeContest = 'https://vjudge.net/contest/';
  const userSchema = ['name', 'id', 'email', 'vjudge', 'codeforces', 'atcoder'];

  const users = getValuesFromInputSheet<User[]>(
    tableHeadingInSheet.user,
    [],
    (previousValue, currentRow, currentColumn, sheet) => {
      const values = sheet
        .getRange(currentRow, currentColumn, 1, 6)
        .getValues()[0] as string[];
      previousValue.push({
        name: values[0],
        id: values[1].toString(),
        email: values[2],
        handles: {
          vjudge: values[3].split(',').map(s => s.trim()),
          codeforces: values[4].split(',').map(s => s.trim()),
          atcoder: values[5].split(',').map(s => s.trim())
        }
      })
      return previousValue.filter(user => user.id && user.handles.vjudge);
    }
  );

  const vjudgeContests = getValuesFromInputSheet<VjudgeContest[]>(
    tableHeadingInSheet.contest,
    [],
    (previousValue, currentRow, currentColumn, sheet) => {
      const values = sheet
        .getRange(currentRow, currentColumn, 1, 3)
        .getValues()[0] as string[];
      previousValue.push({
        id: values[0],
        reqCount: parseInt(values[1]),
        reqProblems: Array.from(
          new Set(values[2].split(',').map(id => id.trim().toUpperCase()))
        )
          .filter(s => 0 < s.length && s.length <= 2 && /^[A-Z]+$/.test(s))
          .sort((a, b) => probIdToIndex(a) - probIdToIndex(b))
      })
      return previousValue;
    }
  );

  const colorHex = getValuesFromInputSheet<Partial<Record<color, string>>>(
    tableHeadingInSheet.color,
    {},
    (previousValue, currentRow, currentColumn, sheet) => {
      const colorName = sheet
        .getRange(currentRow, currentColumn)
        .getValue() as string;
      const colorHex = sheet
        .getRange(currentRow, currentColumn + 1)
        .getBackground();
      return {
        ...previousValue,
        [colorName]: colorHex
      };
    }
  ) as Record<color, string>;

  const solveMap: Record<string, Record<string, Set<string>>> = {};
  const solveCountMap: Record<string, Record<string, number>> = {};

  for (const contest of vjudgeContests) {
    const url = baseUrlVjudgeContest + 'rank/single/' + contest.id;
    const respone = UrlFetchApp.fetch(url);
    const data = JSON.parse(respone.getContentText()) as VjudgeResponseData;

    contest.title = data.title;
    const vjudgeIdHandleMap = Object.entries(data.participants)
      .reduce(
        (result, [key, value]) => ({
          ...result,
          [key]: value[0]
        }),
        {}
      );

    const submissions = data.submissions
      .filter(
        sub => (
          sub[2] === 1 &&
          null !== findUserId(vjudgeIdHandleMap[sub[0].toString()], 'vjudge')
        )
      )
      .map(
        sub => ({
          userId: findUserId(vjudgeIdHandleMap[sub[0].toString()], 'vjudge'),
          problemIndex: sub[1],
          time: sub[3]
        })
      );

    submissions.forEach(
      sub => {
        if (!solveMap[sub.userId]) {
          solveMap[sub.userId] = {};
        }
        if (!solveMap[sub.userId][contest.id]) {
          solveMap[sub.userId][contest.id] = new Set();
        }
        solveMap[sub.userId][contest.id].add(probIndexToId(sub.problemIndex))
      }
    );

    users.forEach(
      user => {
        vjudgeContests.forEach(
          contest => {
            if (!solveMap[user.id]) {
              solveMap[user.id] = {};
            }
            if (!solveMap[user.id][contest.id]) {
              solveMap[user.id][contest.id] = new Set();
            }
            const solves = solveMap[user.id][contest.id];
            if (!solveCountMap[user.id]) {
              solveCountMap[user.id] = {};
            }
            solveCountMap[user.id][contest.id] =
              Math.min(
                solves.size,
                contest.reqCount - contest.reqProblems.reduce(
                  (count, problemId) => count + (solves.has(problemId) ? 0 : 1),
                  0
                )
              );
          }
        )
      }
    )
  }

  const totalProblems = vjudgeContests.reduce(
    (total, contest) => total + contest.reqCount,
    0
  );

  // write to sheet
  (() => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(sheetName.outpuSolveTracker);
    const previousRange = sheet.getDataRange();
    previousRange.breakApart();

    const formatRules = sheet.getConditionalFormatRules();
    while (formatRules.length) formatRules.pop();

    let column = 1;

    // set rank heading
    sheet.setColumnWidth(column, 20);
    sheet.getRange(1, column, 2, 1)
      .merge()
      .setValue('Rank')
      .setWrap(true)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    ++column;

    // set user ("participant") name heading
    sheet.setColumnWidth(column, 200);
    sheet.getRange(1, column, 2, 1)
      .merge()
      .setValue('Participant name')
      .setWrap(true)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    ++column;

    // set total solves heading
    sheet.getRange(1, column, 1, 2)
      .merge()
      .setValue('Total solves')
      .setWrap(true)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    // set total required solve-count sub-heading
    sheet.setColumnWidth(column, 90);
    sheet.getRange(2, column)
      .setValue(`Required: ${totalProblems}`);

    // build and push format rules for total solve-count
    const rulesTotalSolves = formatRulesForSolveCount(
      totalProblems,
      sheet.getRange(3, column, users.length, 1)
    );
    rulesTotalSolves.forEach(rule => formatRules.push(rule));
    ++column;

    // set total reqruied solve percent sub-heading
    sheet.setColumnWidth(column, 40);
    sheet.getRange(2, column)
      .setValue(`100%`);

    // build and push format rule for total solve percent
    const rulePercentGradient = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue(
        colorHex.green,
        SpreadsheetApp.InterpolationType.PERCENT,
        '100'
      )
      .setGradientMidpointWithValue(
        colorHex.yellow,
        SpreadsheetApp.InterpolationType.PERCENT,
        '75'
      )
      .setGradientMinpointWithValue(
        colorHex.red,
        SpreadsheetApp.InterpolationType.PERCENT,
        '50'
      )
      .setRanges([sheet.getRange(3, column, users.length, 1)])
      .build();
    formatRules.push(rulePercentGradient);
    ++column;

    // set user credentials headings
    [
      'NSU ID',
      'Email',
      'VJudge handle(s)',
      'Codeforces handle(s)',
      'AtCoder handle(s)'
    ].forEach(
      text => {
        sheet.setColumnWidth(column, 150)
        sheet.getRange(1, column, 2, 1)
          .merge()
          .setValue(text)
          .setWrap(true)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        ++column
      }
    );

    // loop over contests to set their headings and sub-headings
    vjudgeContests.forEach(
      contest => {
        // build rich-text for contest's heading
        const richtTextContestHeading = SpreadsheetApp.newRichTextValue()
          .setText(contest.title)
          .setLinkUrl(baseUrlVjudgeContest + contest.id)
          .build();

        // set contest's heading
        if (contest.reqProblems.length) {
          sheet.getRange(1, column, 1, contest.reqProblems.length + 1)
            .merge()
        }
        sheet.getRange(1, column, 1, contest.reqProblems.length + 1)
          .setValue('')
          .setRichTextValue(richtTextContestHeading)
          .setWrap(true)
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

        // set contest's required solve-count sub-heading
        sheet.setColumnWidth(column, 90);
        sheet.getRange(2, column)
          .setValue(`Required: ${contest.reqCount}`);

        // build and push format rule for solve-count
        const rulesContestSolves = formatRulesForSolveCount(
          contest.reqCount,
          sheet.getRange(3, column, users.length, 1)
        );
        rulesContestSolves.forEach(rule => formatRules.push(rule));
        ++column

        // set contest's required problems sub-heading
        contest.reqProblems.forEach(
          probId => {
            sheet.setColumnWidth(column, 20);
            sheet.getRange(2, column)
              .setValue(probId)
              .setHorizontalAlignment('center');
            ++column;
          }
        );
      }
    );

    sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .setHorizontalAlignment('center')
      .setVerticalAlignment('top');

    sheet.getRange(2, 1, 1, sheet.getLastColumn())
      .setHorizontalAlignment('right');

    sheet.getRange(1, 1, users.length + 2, 4)
      .setBackground('#d0e0e3');

    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(4);

    users
      .map(
        user => ({
          user,
          totalSolves: vjudgeContests.reduce(
            (result, contest) => result + solveCountMap[user.id][contest.id],
            0
          )
        })
      )
      .sort((a, b) => b.totalSolves - a.totalSolves)
      .forEach(
        ({ user, totalSolves }, userIndex) => {
          const userRow = 3 + userIndex;

          let column = 1;

          // set user's rank
          sheet.getRange(userRow, column)
            .setValue(userIndex + 1)
            .setWrap(true)
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
          ++column;

          // set user's name
          sheet.getRange(userRow, column)
            .setValue(user.name)
            .setWrap(true)
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
          ++column;

          // set user's total solve count
          sheet.getRange(userRow, column)
            .setValue(totalSolves)
          ++column;

          // set user's total solve percentage
          sheet.getRange(userRow, column)
            .setValue(`${Math.floor((totalSolves * 100) / totalProblems)}%`)
          ++column;

          // set user's credentials
          [
            user.id,
            user.email,
            user.handles.vjudge,
            user.handles.codeforces,
            user.handles.atcoder
          ].forEach(h => {
            sheet.getRange(userRow, column)
              .setValue(typeof h === 'string' ? h : (h && h.join ? h.join(',') : ''))
              .setWrap(true)
              .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
            ++column
          });

          // set user's contest solves
          vjudgeContests.forEach(
            contest => {
              // set user's solve-count in current contest
              sheet.getRange(userRow, column)
                .setValue(solveCountMap[user.id][contest.id]);
              ++column;

              // set user's solve status in current contest's required problems
              contest.reqProblems.forEach(
                p => {
                  sheet.getRange(userRow, column)
                    .setValue(solveMap[user.id][contest.id].has(p) ? '✅' : '❌')
                    .setHorizontalAlignment('center');
                  ++column;
                }
              );
            }
          );
        }
      );
      
    // clear extra ranges left over from previous script
    const numExtraRows = previousRange.getLastRow() - users.length;
    if (numExtraRows > 0) {
      sheet.getRange(
        users.length + 1,
        1,
        numExtraRows,
        sheet.getLastColumn()
      )
        .clear();
    }
    const numExtraColumns = previousRange.getLastColumn() - (column - 1);
    if (numExtraColumns > 0) {
      sheet.getRange(
        1,
        column,
        sheet.getLastRow(),
        numExtraColumns
      )
        .clear();
    }

    sheet.setConditionalFormatRules(formatRules);
  })();

  function getValuesFromInputSheet<T>(
    tableHeadingInSheet: string,
    initialValue: T,
    addValueFn: (
      previousValue: T,
      currentRow: number,
      currentColumn: number,
      sheet: GoogleAppsScript.Spreadsheet.Sheet
    ) => T
  ) {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet().getSheetByName(sheetName.input);
    const head = sheet.createTextFinder(tableHeadingInSheet).findAll()[0];
    const column = head.getColumn();
    let result = initialValue;
    let range = sheet.getRange(head.getRow() + 2, column);
    while (range.getValue() as string != '') {
      result = addValueFn(result, range.getRow(), range.getColumn(), sheet);
      range = sheet.getRange(range.getRow() + 1, column);
    }
    return result;
  }

  function findUserId(handle: string, judge: Judge): string {
    if (!handle) return null;
    for (const user of users) {
      if (!!user.handles[judge].find(h => handle === h)) {
        return user.id;
      }
    }
    return null;
  }

  function probIdToIndex(id: string): number {
    if (id.length > 2) throw 'cannot convert id to index';
    return id.toUpperCase().split('').reduce(
      (res, ch) => (res * 26) + (ch.charCodeAt(0) - 'A'.charCodeAt(0) + 1),
      0
    ) - 1;
  }

  function probIndexToId(index: number): string {
    if (index < 0 || 676 <= index) throw 'cannot convert index to id';
    if (index < 26) {
      return String.fromCharCode('A'.charCodeAt(0) + index);
    }
    let res: string[] = [];
    while (index >= 26) {
      const d = index % 26;
      res.unshift(String.fromCharCode('A'.charCodeAt(0) + d));
      index = Math.floor(index / 26);
    }
    res.unshift(String.fromCharCode('A'.charCodeAt(0) + index - 1));
    return res.join('').toUpperCase();
  }

  function formatRulesForSolveCount(
    targetSolves: number,
    range: GoogleAppsScript.Spreadsheet.Range
  ): GoogleAppsScript.Spreadsheet.ConditionalFormatRule[] {
    const rule100 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(targetSolves)
      .setBackground(colorHex.lightGreen)
      .setRanges([range])
      .build();
    const rule066 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(Math.ceil(targetSolves * 0.66))
      .setBackground(colorHex.lightYellow)
      .setRanges([range])
      .build();
    const rule033 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(Math.ceil(targetSolves * 0.33))
      .setBackground(colorHex.lightOrange)
      .setRanges([range])
      .build();
    const rule000 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0)
      .setBackground(colorHex.lightRed)
      .setRanges([range])
      .build();
    return [rule100, rule066, rule033, rule000];
  }
}