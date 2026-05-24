function buildPlayerMatchLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('matches_raw');
  const out = ss.getSheetByName('player_match_log');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!out) throw new Error('No existe la hoja player_match_log');

  const data = raw.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const rows = data.slice(1);

  const idx = indexMap(headers);
  const orderedRows = sortRowsByPlayedOrder(rows, idx);

  const outputHeaders = [
    'log_id','match_id','campeonato','fecha','division',
    'player','partner','opp_1','opp_2','side',
    'won_match','lost_match','match_result',
    'sets_won','sets_lost','games_won','games_lost',
    'match_format','won_2s','lost_2s','won_3s','lost_3s',
    'cbs','farra',
    'set1_result','set2_result','set3_result',
    'set1_type','set2_type','set3_type',
    'ranking_points','is_counted_for_last14','chronological_index'
  ];

  const outRows = [];
  let chrono = 1;

  orderedRows.forEach((r) => {
    const matchId = val(r, idx, 'match_id');
    const campeonato = val(r, idx, 'campeonato');
    const fecha = val(r, idx, 'fecha');
    const division = val(r, idx, 'division');

    const p1j1 = val(r, idx, 'pareja1_j1');
    const p1j2 = val(r, idx, 'pareja1_j2');
    const p2j1 = val(r, idx, 'pareja2_j1');
    const p2j2 = val(r, idx, 'pareja2_j2');

    const winner = clean(val(r, idx, 'winner'));
    const s1p1 = num(val(r, idx, 'set1_p1'));
    const s1p2 = num(val(r, idx, 'set1_p2'));
    const s2p1 = num(val(r, idx, 'set2_p1'));
    const s2p2 = num(val(r, idx, 'set2_p2'));
    const s3p1 = numOrBlank(val(r, idx, 'set3_p1'));
    const s3p2 = numOrBlank(val(r, idx, 'set3_p2'));

    const played = isPlayedRow(winner, s1p1, s1p2, s2p1, s2p2);
    if (!played) return;

    const hasSet3 = s3p1 !== '' && s3p2 !== '';
    const format = hasSet3 ? '3S' : '2S';

    const p1 = [p1j1, p1j2];
    const p2 = [p2j1, p2j2];

    const setsSummaryP1 = getSetSummary(
      [
        [s1p1, s1p2],
        [s2p1, s2p2],
        hasSet3 ? [s3p1, s3p2] : null
      ],
      'P1'
    );

    const setsSummaryP2 = getSetSummary(
      [
        [s1p1, s1p2],
        [s2p1, s2p2],
        hasSet3 ? [s3p1, s3p2] : null
      ],
      'P2'
    );

    const p1Won = winner === 'P1';
    const p2Won = winner === 'P2';

    outRows.push(buildLogRow(matchId, campeonato, fecha, division, p1j1, p1j2, p2j1, p2j2, 'P1', p1Won, format, setsSummaryP1, chrono));
    outRows.push(buildLogRow(matchId, campeonato, fecha, division, p1j2, p1j1, p2j1, p2j2, 'P1', p1Won, format, setsSummaryP1, chrono));
    outRows.push(buildLogRow(matchId, campeonato, fecha, division, p2j1, p2j2, p1j1, p1j2, 'P2', p2Won, format, setsSummaryP2, chrono));
    outRows.push(buildLogRow(matchId, campeonato, fecha, division, p2j2, p2j1, p1j1, p1j2, 'P2', p2Won, format, setsSummaryP2, chrono));

    chrono++;
  });

  out.clearContents();
  out.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, outputHeaders.length).setValues(outRows);
  }
}

function buildLogRow(matchId, campeonato, fecha, division, player, partner, opp1, opp2, side, wonMatch, format, summary, chrono) {
  const lostMatch = wonMatch ? 0 : 1;
  const result = wonMatch ? 'W' : 'L';

  const won2s = wonMatch && format === '2S' ? 1 : 0;
  const lost2s = !wonMatch && format === '2S' ? 1 : 0;
  const won3s = wonMatch && format === '3S' ? 1 : 0;
  const lost3s = !wonMatch && format === '3S' ? 1 : 0;

  const cbs = (wonMatch && format === '3S' && summary.set1Result === 'L' && summary.set2Result === 'W' && summary.set3Result === 'W') ? 1 : 0;
  const farra = (!wonMatch && format === '3S' && summary.set1Result === 'W' && summary.set2Result === 'L' && summary.set3Result === 'L') ? 1 : 0;

  const rankingPoints = ''; // lo dejamos vacío por ahora

  return [
    `${matchId}_${player}`,
    matchId,
    campeonato,
    fecha,
    division,
    player,
    partner,
    opp1,
    opp2,
    side,
    wonMatch ? 1 : 0,
    lostMatch,
    result,
    summary.setsWon,
    summary.setsLost,
    summary.gamesWon,
    summary.gamesLost,
    format,
    won2s,
    lost2s,
    won3s,
    lost3s,
    cbs,
    farra,
    summary.set1Result,
    summary.set2Result,
    summary.set3Result,
    summary.set1Type,
    summary.set2Type,
    summary.set3Type,
    rankingPoints,
    '',
    chrono
  ];
}

function getSetSummary(sets, side) {
  let setsWon = 0;
  let setsLost = 0;
  let gamesWon = 0;
  let gamesLost = 0;

  const [set1, set2, set3] = sets;

  const parsed1 = parseSet(set1, side);
  const parsed2 = parseSet(set2, side);
  const parsed3 = set3 ? parseSet(set3, side) : emptySet();

  [parsed1, parsed2, parsed3].forEach((s) => {
    if (!s.exists) return;
    gamesWon += s.my;
    gamesLost += s.opp;
    if (s.result === 'W') setsWon++;
    if (s.result === 'L') setsLost++;
  });

  return {
    setsWon,
    setsLost,
    gamesWon,
    gamesLost,
    set1Result: parsed1.result,
    set2Result: parsed2.result,
    set3Result: parsed3.result,
    set1Type: parsed1.type,
    set2Type: parsed2.type,
    set3Type: parsed3.type
  };
}

function parseSet(setPair, side) {
  if (!setPair || setPair[0] === '' || setPair[1] === '') return emptySet();

  const p1 = Number(setPair[0]);
  const p2 = Number(setPair[1]);

  const my = side === 'P1' ? p1 : p2;
  const opp = side === 'P1' ? p2 : p1;

  return {
    exists: true,
    my,
    opp,
    result: my > opp ? 'W' : 'L',
    type: classifySet(my, opp)
  };
}

function classifySet(my, opp) {
  const hi = Math.max(my, opp);
  const lo = Math.min(my, opp);

  if (hi === 7 && lo === 6) return 'TB';
  if ((hi === 6 && lo === 4) || (hi === 7 && lo === 5)) return 'TIGHT';
  if (hi === 6 && (lo === 2 || lo === 3)) return 'COLD';
  if (hi === 6 && (lo === 0 || lo === 1)) return 'HAMMER';
  return '';
}

function emptySet() {
  return {
    exists: false,
    my: 0,
    opp: 0,
    result: '',
    type: ''
  };
}

function isPlayedRow(winner, s1p1, s1p2, s2p1, s2p2) {
  return (winner === 'P1' || winner === 'P2') &&
         s1p1 !== '' && s1p2 !== '' &&
         s2p1 !== '' && s2p2 !== '';
}

function indexMap(headers) {
  const map = {};
  headers.forEach((h, i) => map[String(h).trim()] = i);
  return map;
}

function val(row, idx, key) {
  return row[idx[key]];
}

function clean(v) {
  return String(v || '').trim();
}

function num(v) {
  if (v === '' || v === null || typeof v === 'undefined') return '';
  return Number(v);
}

function numOrBlank(v) {
  if (v === '' || v === null || typeof v === 'undefined') return '';
  return Number(v);
}
function buildH2HPairs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('matches_raw');
  const out = ss.getSheetByName('h2h_pairs');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!out) throw new Error('No existe la hoja h2h_pairs');

  const data = raw.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const rows = data.slice(1);
  const idx = indexMap(headers);
  const orderedRows = sortRowsByPlayedOrder(rows, idx);

  const outputHeaders = [
    'h2h_id','player_a','player_b',
    'matches_a','matches_b','matches_wr_a','matches_wr_b',
    'sets_a','sets_b','sets_wr_a','sets_wr_b',
    'games_a','games_b','games_wr_a','games_wr_b',
    'last_match_id','last_match_camp','last_match_fecha','history_count'
  ];

  const map = {};

  orderedRows.forEach((r) => {
    const matchId = val(r, idx, 'match_id');
    const campeonato = val(r, idx, 'campeonato');
    const fecha = val(r, idx, 'fecha');
    const division = val(r, idx, 'division');

    const p1j1 = clean(val(r, idx, 'pareja1_j1'));
    const p1j2 = clean(val(r, idx, 'pareja1_j2'));
    const p2j1 = clean(val(r, idx, 'pareja2_j1'));
    const p2j2 = clean(val(r, idx, 'pareja2_j2'));

    const winner = clean(val(r, idx, 'winner'));
    const s1p1 = num(val(r, idx, 'set1_p1'));
    const s1p2 = num(val(r, idx, 'set1_p2'));
    const s2p1 = num(val(r, idx, 'set2_p1'));
    const s2p2 = num(val(r, idx, 'set2_p2'));
    const s3p1 = numOrBlank(val(r, idx, 'set3_p1'));
    const s3p2 = numOrBlank(val(r, idx, 'set3_p2'));

    const played = isPlayedRow(winner, s1p1, s1p2, s2p1, s2p2);
    if (!played) return;

    const setPairs = [
      [s1p1, s1p2],
      [s2p1, s2p2]
    ];
    if (s3p1 !== '' && s3p2 !== '') {
      setPairs.push([s3p1, s3p2]);
    }

    const leftPlayers = [p1j1, p1j2];
    const rightPlayers = [p2j1, p2j2];

    leftPlayers.forEach((leftPlayer) => {
      rightPlayers.forEach((rightPlayer) => {
        const pair = [leftPlayer, rightPlayer].sort();
        const playerA = pair[0];
        const playerB = pair[1];
        const h2hId = `${playerA}|${playerB}`;

        if (!map[h2hId]) {
          map[h2hId] = {
            h2h_id: h2hId,
            player_a: playerA,
            player_b: playerB,
            matches_a: 0,
            matches_b: 0,
            sets_a: 0,
            sets_b: 0,
            games_a: 0,
            games_b: 0,
            last_match_id: '',
            last_match_camp: '',
            last_match_fecha: '',
            history_count: 0
          };
        }

        const obj = map[h2hId];

        const aSide = leftPlayer === playerA ? 'P1' : 'P2';
        const bSide = aSide === 'P1' ? 'P2' : 'P1';

        // Match winner
        if (winner === aSide) obj.matches_a += 1;
        if (winner === bSide) obj.matches_b += 1;

        // Sets and games
        setPairs.forEach(([p1, p2]) => {
          const aGames = aSide === 'P1' ? p1 : p2;
          const bGames = aSide === 'P1' ? p2 : p1;

          obj.games_a += aGames;
          obj.games_b += bGames;

          if (aGames > bGames) obj.sets_a += 1;
          if (bGames > aGames) obj.sets_b += 1;
        });

        obj.last_match_id = matchId;
        obj.last_match_camp = campeonato;
        obj.last_match_fecha = fecha;
        obj.history_count += 1;
      });
    });
  });

  const outRows = Object.values(map).map((obj) => {
    const totalMatches = obj.matches_a + obj.matches_b;
    const totalSets = obj.sets_a + obj.sets_b;
    const totalGames = obj.games_a + obj.games_b;

    const matchesWrA = totalMatches > 0 ? obj.matches_a / totalMatches : '';
    const matchesWrB = totalMatches > 0 ? obj.matches_b / totalMatches : '';

    const setsWrA = totalSets > 0 ? obj.sets_a / totalSets : '';
    const setsWrB = totalSets > 0 ? obj.sets_b / totalSets : '';

    const gamesWrA = totalGames > 0 ? obj.games_a / totalGames : '';
    const gamesWrB = totalGames > 0 ? obj.games_b / totalGames : '';

    return [
      obj.h2h_id,
      obj.player_a,
      obj.player_b,
      obj.matches_a,
      obj.matches_b,
      matchesWrA,
      matchesWrB,
      obj.sets_a,
      obj.sets_b,
      setsWrA,
      setsWrB,
      obj.games_a,
      obj.games_b,
      gamesWrA,
      gamesWrB,
      obj.last_match_id,
      obj.last_match_camp,
      obj.last_match_fecha,
      obj.history_count
    ];
  });

  outRows.sort((a, b) => {
    if (a[1] < b[1]) return -1;
    if (a[1] > b[1]) return 1;
    if (a[2] < b[2]) return -1;
    if (a[2] > b[2]) return 1;
    return 0;
  });

  out.clearContents();
  out.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, outputHeaders.length).setValues(outRows);

    // Apply percentage formats
    out.getRange(2, 6, outRows.length, 2).setNumberFormat('0.0%');
    out.getRange(2, 10, outRows.length, 2).setNumberFormat('0.0%');
    out.getRange(2, 14, outRows.length, 2).setNumberFormat('0.0%');
  }
}
function buildH2HPlayerView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName('h2h_pairs');
  const out = ss.getSheetByName('h2h_player_view');

  if (!base) throw new Error('No existe la hoja h2h_pairs');
  if (!out) throw new Error('No existe la hoja h2h_player_view');

  const data = base.getDataRange().getValues();
  if (data.length < 2) throw new Error('h2h_pairs no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const requiredHeaders = [
    'h2h_id','player_a','player_b',
    'matches_a','matches_b','matches_wr_a','matches_wr_b',
    'sets_a','sets_b','sets_wr_a','sets_wr_b',
    'games_a','games_b','games_wr_a','games_wr_b',
    'last_match_id','last_match_camp','last_match_fecha','history_count'
  ];

  const missing = requiredHeaders.filter(h => !headers.includes(h));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en h2h_pairs: ' + missing.join(', '));
  }

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const outputHeaders = [...requiredHeaders];
  const outRows = [];

  rows.forEach((r) => {
    const h2hId = r[idx['h2h_id']];
    const playerA = r[idx['player_a']];
    const playerB = r[idx['player_b']];

    const matchesA = r[idx['matches_a']];
    const matchesB = r[idx['matches_b']];
    const matchesWrA = r[idx['matches_wr_a']];
    const matchesWrB = r[idx['matches_wr_b']];

    const setsA = r[idx['sets_a']];
    const setsB = r[idx['sets_b']];
    const setsWrA = r[idx['sets_wr_a']];
    const setsWrB = r[idx['sets_wr_b']];

    const gamesA = r[idx['games_a']];
    const gamesB = r[idx['games_b']];
    const gamesWrA = r[idx['games_wr_a']];
    const gamesWrB = r[idx['games_wr_b']];

    const lastMatchId = r[idx['last_match_id']];
    const lastMatchCamp = r[idx['last_match_camp']];
    const lastMatchFecha = r[idx['last_match_fecha']];
    const historyCount = r[idx['history_count']];

    // fila directa
    outRows.push([
      h2hId,
      playerA,
      playerB,
      matchesA,
      matchesB,
      matchesWrA,
      matchesWrB,
      setsA,
      setsB,
      setsWrA,
      setsWrB,
      gamesA,
      gamesB,
      gamesWrA,
      gamesWrB,
      lastMatchId,
      lastMatchCamp,
      lastMatchFecha,
      historyCount
    ]);

    // fila espejo
    outRows.push([
      h2hId,
      playerB,
      playerA,
      matchesB,
      matchesA,
      matchesWrB,
      matchesWrA,
      setsB,
      setsA,
      setsWrB,
      setsWrA,
      gamesB,
      gamesA,
      gamesWrB,
      gamesWrA,
      lastMatchId,
      lastMatchCamp,
      lastMatchFecha,
      historyCount
    ]);
  });

  outRows.sort((a, b) => {
    if (a[1] < b[1]) return -1;
    if (a[1] > b[1]) return 1;
    if (a[2] < b[2]) return -1;
    if (a[2] > b[2]) return 1;
    return 0;
  });

  out.clearContents();
  out.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, outputHeaders.length).setValues(outRows);

    // formato porcentaje
    out.getRange(2, 6, outRows.length, 2).setNumberFormat('0.0%');
    out.getRange(2, 10, outRows.length, 2).setNumberFormat('0.0%');
    out.getRange(2, 14, outRows.length, 2).setNumberFormat('0.0%');
  }

  notifyUser(
    'h2h_player_view reconstruida\\n' +
    'Filas generadas: ' + outRows.length
  );
}
function buildH2HMatchLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('matches_raw');
  const out = ss.getSheetByName('h2h_match_log');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!out) throw new Error('No existe la hoja h2h_match_log');

  const data = raw.getDataRange().getValues();
  if (data.length < 2) throw new Error('matches_raw no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  const orderedRows = sortRowsByPlayedOrder(rows, idx);

  const requiredHeaders = [
    'match_id', 'campeonato', 'fecha', 'division',
    'pareja1_j1', 'pareja1_j2', 'pareja2_j1', 'pareja2_j2',
    'winner', 'set1_p1', 'set1_p2', 'set2_p1', 'set2_p2', 'set3_p1', 'set3_p2'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en matches_raw: ' + missing.join(', '));
  }

  const outputHeaders = [
    'h2h_id', 'match_id', 'campeonato', 'fecha', 'division',
    'player_a', 'player_b', 'player_a_side', 'player_b_side',
    'pair_a_display', 'pair_b_display', 'score_display',
    'winner_pair', 'winner_h2h',
    'matches_a_increment', 'matches_b_increment',
    'sets_a', 'sets_b', 'games_a', 'games_b',
    'chronological_index'
  ];

  const outRows = [];
  let chronologicalIndex = 1;

  for (let i = 0; i < orderedRows.length; i++) {
  const r = orderedRows[i];


    const matchId = safe(r[idx['match_id']]);
    const campeonato = safe(r[idx['campeonato']]);
    const fecha = safe(r[idx['fecha']]);
    const division = safe(r[idx['division']]);

    const p1j1 = safe(r[idx['pareja1_j1']]);
    const p1j2 = safe(r[idx['pareja1_j2']]);
    const p2j1 = safe(r[idx['pareja2_j1']]);
    const p2j2 = safe(r[idx['pareja2_j2']]);

    const winner = safe(r[idx['winner']]);
    const s1p1 = r[idx['set1_p1']];
    const s1p2 = r[idx['set1_p2']];
    const s2p1 = r[idx['set2_p1']];
    const s2p2 = r[idx['set2_p2']];
    const s3p1 = r[idx['set3_p1']];
    const s3p2 = r[idx['set3_p2']];

    const played =
      (winner === 'P1' || winner === 'P2') &&
      s1p1 !== '' && s1p1 !== null &&
      s1p2 !== '' && s1p2 !== null &&
      s2p1 !== '' && s2p1 !== null &&
      s2p2 !== '' && s2p2 !== null;

    if (!played) continue;

    const leftPlayers = [p1j1, p1j2];
    const rightPlayers = [p2j1, p2j2];

    const setPairs = [
      [Number(s1p1), Number(s1p2)],
      [Number(s2p1), Number(s2p2)]
    ];

    if (s3p1 !== '' && s3p1 !== null && s3p2 !== '' && s3p2 !== null) {
      setPairs.push([Number(s3p1), Number(s3p2)]);
    }

    const scoreDisplay = setPairs.map(([a, b]) => `${a}-${b}`).join(', ');
    const pairLeft = `${p1j1} / ${p1j2}`;
    const pairRight = `${p2j1} / ${p2j2}`;
    const winnerPair = winner === 'P1' ? pairLeft : pairRight;

    // 4 cruces rivales
    const rivalPairs = [
      [p1j1, p2j1],
      [p1j1, p2j2],
      [p1j2, p2j1],
      [p1j2, p2j2]
    ];

    for (let j = 0; j < rivalPairs.length; j++) {
      const leftPlayer = rivalPairs[j][0];
      const rightPlayer = rivalPairs[j][1];

      if (!leftPlayer || !rightPlayer) continue;

      const h2hId = [leftPlayer, rightPlayer].sort().join('|');

      let setsA = 0, setsB = 0, gamesA = 0, gamesB = 0;

      for (let k = 0; k < setPairs.length; k++) {
        const a = setPairs[k][0];
        const b = setPairs[k][1];
        gamesA += a;
        gamesB += b;
        if (a > b) setsA++;
        else if (b > a) setsB++;
      }

      // fila directa
      outRows.push([
        h2hId,
        matchId,
        campeonato,
        fecha,
        division,
        leftPlayer,
        rightPlayer,
        'P1',
        'P2',
        pairLeft,
        pairRight,
        scoreDisplay,
        winnerPair,
        winner === 'P1' ? leftPlayer : rightPlayer,
        winner === 'P1' ? 1 : 0,
        winner === 'P2' ? 1 : 0,
        setsA,
        setsB,
        gamesA,
        gamesB,
        chronologicalIndex
      ]);

      // fila espejo
      outRows.push([
        h2hId,
        matchId,
        campeonato,
        fecha,
        division,
        rightPlayer,
        leftPlayer,
        'P2',
        'P1',
        pairRight,
        pairLeft,
        scoreDisplay,
        winnerPair,
        winner === 'P2' ? rightPlayer : leftPlayer,
        winner === 'P2' ? 1 : 0,
        winner === 'P1' ? 1 : 0,
        setsB,
        setsA,
        gamesB,
        gamesA,
        chronologicalIndex
      ]);
    }

    chronologicalIndex++;
  }

  outRows.sort((a, b) => {
    if (a[5] < b[5]) return -1;
    if (a[5] > b[5]) return 1;
    if (a[6] < b[6]) return -1;
    if (a[6] > b[6]) return 1;
    return b[20] - a[20];
  });

  out.clearContents();
  out.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, outputHeaders.length).setValues(outRows);
  }

  notifyUser(
    'h2h_match_log reconstruida\n' +
    'Filas generadas: ' + outRows.length
  );
}
function safe(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}
function buildRankingLast14() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('matches_raw');
  const out = ss.getSheetByName('ranking_last14');
  const currentStatus = ss.getSheetByName('current_player_status');
if (!currentStatus) throw new Error('No existe la hoja current_player_status');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!out) throw new Error('No existe la hoja ranking_last14');

  const data = raw.getDataRange().getValues();
  if (data.length < 2) throw new Error('matches_raw no tiene datos');
  const statusData = currentStatus.getDataRange().getValues();
if (statusData.length < 2) throw new Error('current_player_status no tiene datos');

const statusHeaders = statusData[0].map(h => String(h).trim());
const statusRows = statusData.slice(1);

const statusIdx = {};
statusHeaders.forEach((h, i) => statusIdx[h] = i);

if (!('player' in statusIdx) || !('division_actual' in statusIdx)) {
  throw new Error('current_player_status debe tener columnas player y division_actual');
}

const currentDivisionMap = {};
statusRows.forEach(r => {
  const player = safe(r[statusIdx['player']]);
  if (!player) return;
  currentDivisionMap[player] = safe(r[statusIdx['division_actual']]) || '-';
});

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  const orderedRows = sortRowsByPlayedOrder(rows, idx);

  const requiredHeaders = [
    'campeonato', 'fecha', 'division',
    'pareja1_j1', 'pareja1_j2', 'pareja2_j1', 'pareja2_j2',
    'winner', 'set1_p1', 'set1_p2', 'set2_p1', 'set2_p2', 'set3_p1', 'set3_p2'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en matches_raw: ' + missing.join(', '));
  }

  const outputHeaders = [
  'rank_position','player','division_actual','matches_played','matches_counted',
  'score_last14','avg_last14',
  'count_div_a','count_div_b','count_div_c','count_div_d',
  'g2','g3','p3','p2',
  'record_last14','result_mix_last14'
];

  const playerMatches = {};
  const lastDivision = {};

  for (let i = 0; i < orderedRows.length; i++) {
  const r = orderedRows[i];

    const division = safe(r[idx['division']]);
    const p1a = safe(r[idx['pareja1_j1']]);
    const p1b = safe(r[idx['pareja1_j2']]);
    const p2a = safe(r[idx['pareja2_j1']]);
    const p2b = safe(r[idx['pareja2_j2']]);

    const winner = safe(r[idx['winner']]);
    const s1a = toNumOrZero(r[idx['set1_p1']]);
    const s1b = toNumOrZero(r[idx['set1_p2']]);
    const s2a = toNumOrZero(r[idx['set2_p1']]);
    const s2b = toNumOrZero(r[idx['set2_p2']]);
    const s3aRaw = r[idx['set3_p1']];
    const s3bRaw = r[idx['set3_p2']];
    const s3a = toNumOrZero(s3aRaw);
    const s3b = toNumOrZero(s3bRaw);

    const played =
      division !== '' &&
      p1a !== '' && p1b !== '' && p2a !== '' && p2b !== '' &&
      (winner === 'P1' || winner === 'P2') &&
      r[idx['set1_p1']] !== '' && r[idx['set1_p1']] !== null &&
      r[idx['set1_p2']] !== '' && r[idx['set1_p2']] !== null &&
      r[idx['set2_p1']] !== '' && r[idx['set2_p1']] !== null &&
      r[idx['set2_p2']] !== '' && r[idx['set2_p2']] !== null;

    if (!played) continue;

    let sets1 = 0, sets2 = 0;
    const games1 = s1a + s2a + s3a;
    const games2 = s1b + s2b + s3b;

    if (s1a > s1b) sets1++; else sets2++;
    if (s2a > s2b) sets1++; else sets2++;

    if ((s3a !== 0 || s3b !== 0)) {
      if (s3a > s3b) sets1++;
      else sets2++;
    }

    let res1 = '', res2 = '';

    if (winner === 'P1') {
      if (sets1 === 2 && sets2 === 0) {
        res1 = 'G2'; res2 = 'P2';
      } else {
        res1 = 'G3'; res2 = 'P3';
      }
    } else {
      if (sets2 === 2 && sets1 === 0) {
        res1 = 'P2'; res2 = 'G2';
      } else {
        res1 = 'P3'; res2 = 'G3';
      }
    }

    const pts1 = rankPointsOfficial(division, res1, games1 - games2);
    const pts2 = rankPointsOfficial(division, res2, games2 - games1);

    addPlayerRankingMatch(playerMatches, lastDivision, p1a, division, res1, pts1);
    addPlayerRankingMatch(playerMatches, lastDivision, p1b, division, res1, pts1);
    addPlayerRankingMatch(playerMatches, lastDivision, p2a, division, res2, pts2);
    addPlayerRankingMatch(playerMatches, lastDivision, p2b, division, res2, pts2);
  }

  const players = Object.keys(playerMatches);
  const outRows = [];

  players.forEach(player => {
    const items = playerMatches[player];
    const pj = items.length;
    const pc = Math.min(14, pj);
    const startIdx = pj - pc;

    let score = 0;
    let ca = 0, cb = 0, cc = 0, cd = 0;
    let g2 = 0, g3 = 0, p3 = 0, p2 = 0;

    for (let i = startIdx; i < pj; i++) {
      const item = items[i];
      score += item.points;

      switch (item.division) {
        case 'A': ca++; break;
        case 'B': cb++; break;
        case 'C': cc++; break;
        case 'D': cd++; break;
      }

      switch (item.resultCode) {
        case 'G2': g2++; break;
        case 'G3': g3++; break;
        case 'P3': p3++; break;
        case 'P2': p2++; break;
      }
    }

   const scoreFinal = score;
const avgFinal = pc > 0 ? scoreFinal / pc : 0;

const winsLast14 = g2 + g3;
const lossesLast14 = p3 + p2;
const recordLast14 = `${winsLast14}-${lossesLast14}`;
const resultMixLast14 = `${g2}-${g3}-${p3}-${p2}`;

outRows.push([
  0,
  player,
  currentDivisionMap[player] || '-',
  pj,
  pc,
  round1(scoreFinal),
  round2(avgFinal),
  ca, cb, cc, cd,
  g2, g3, p3, p2,
  recordLast14,
  resultMixLast14
]);
  });

  outRows.sort((a, b) => {
    if (b[5] !== a[5]) return b[5] - a[5];
    return b[6] - a[6];
  });

  for (let i = 0; i < outRows.length; i++) {
    outRows[i][0] = i + 1;
  }

  out.clearContents();

// Fuerza columnas de texto antes de escribir datos
// Columna 16 = record_last14
// Columna 17 = result_mix_last14
out.getRange(1, 16, out.getMaxRows(), 2).setNumberFormat('@');

out.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

if (outRows.length > 0) {
  out.getRange(2, 1, outRows.length, outputHeaders.length).setValues(outRows);
  out.getRange(2, 6, outRows.length, 1).setNumberFormat('0.0');
  out.getRange(2, 7, outRows.length, 1).setNumberFormat('0.00');
  out.getRange(2, 16, outRows.length, 2).setNumberFormat('@');
}

  notifyUser(
    'ranking_last14 reconstruido\n' +
    'Jugadores rankeados: ' + outRows.length
  );
}

function addPlayerRankingMatch(playerMatches, lastDivision, player, division, resultCode, points) {
  if (!player) return;
  if (!playerMatches[player]) playerMatches[player] = [];
  playerMatches[player].push({
    division: division,
    resultCode: resultCode,
    points: points
  });
  lastDivision[player] = division;
}

function rankPointsOfficial(division, resultCode, diffGames) {
  const base = basePointsOfficial(division, resultCode);
  const adj = clamp(
    diffGames * gameMultiplierOfficial(division),
    -gameCapOfficial(division),
    gameCapOfficial(division)
  );

  return base + adj;
}

function basePointsOfficial(division, resultCode) {
  const d = String(division).toUpperCase();
  const map = {
    A: { G2: 1000, G3: 800, P3: 450, P2: 200 },
    B: { G2: 500,  G3: 400, P3: 225, P2: 100 },
    C: { G2: 250,  G3: 200, P3: 113, P2: 50  },
    D: { G2: 125,  G3: 100, P3: 56,  P2: 25  }
  };
  return map[d]?.[resultCode] ?? 0;
}

function floorPointsOfficial(division, resultCode) {
  const d = String(division).toUpperCase();
  const map = {
    A: { G2: 120, G3: 98, P3: 28, P2: 10 },
    B: { G2: 56,  G3: 44, P3: 12, P2: 4  },
    C: { G2: 26,  G3: 20, P3: 5,  P2: 2  },
    D: { G2: 12,  G3: 9,  P3: 2,  P2: 1  }
  };
  return map[d]?.[resultCode] ?? 0;
}

function gameMultiplierOfficial(division) {
  const d = String(division).toUpperCase();
  const map = { A: 12, B: 6, C: 3, D: 1.5 };
  return map[d] ?? 0;
}

function gameCapOfficial(division) {
  const d = String(division).toUpperCase();
  const map = { A: 120, B: 60, C: 30, D: 15 };
  return map[d] ?? 0;
}

function clamp(v, minV, maxV) {
  if (v < minV) return minV;
  if (v > maxV) return maxV;
  return v;
}

function toNumOrZero(v) {
  if (v === '' || v === null || v === undefined) return 0;
  return Number(v);
}

function round1(v) {
  return Math.round(v * 10) / 10;
}

function round2(v) {
  return Math.round(v * 100) / 100;
}
function buildLeagueStatsRankings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('player_match_log');
  const out = ss.getSheetByName('league_stats_rankings');

  if (!log) throw new Error('No existe la hoja player_match_log');
  if (!out) throw new Error('No existe la hoja league_stats_rankings');

  const data = log.getDataRange().getValues();
  if (data.length < 2) throw new Error('player_match_log no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const requiredHeaders = [
    'campeonato',
    'division',
    'player',
    'won_match',
    'lost_match',
    'cbs',
    'farra',
    'set1_result',
    'set2_result',
    'set3_result',
    'set1_type',
    'set2_type',
    'set3_type',
    'chronological_index'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en player_match_log: ' + missing.join(', '));
  }

  const outputHeaders = [
    'league_id',
    'division',
    'category',
    'rank',
    'player',
    'value',
    'label',
    'extra'
  ];

  function safe(v) {
    return v === null || v === undefined ? '' : String(v).trim();
  }

  function num(v) {
    if (v === '' || v === null || typeof v === 'undefined') return 0;
    return Number(v);
  }

  function key(leagueId, division, player) {
    return `${leagueId}|${division}|${player}`;
  }

  function groupKey(leagueId, division) {
    return `${leagueId}|${division}`;
  }

  function divisionWeight(division) {
    const d = String(division || '').trim().toUpperCase();
    if (d === 'A') return 4;
    if (d === 'B') return 3;
    if (d === 'C') return 2;
    if (d === 'D') return 1;
    return 0;
  }

  const playerStats = {};

  function ensurePlayer(leagueId, division, player) {
    const k = key(leagueId, division, player);

    if (!playerStats[k]) {
      playerStats[k] = {
        league_id: leagueId,
        division: division,
        player: player,
        matches: 0,
        wins: 0,
        losses: 0,
        cbs: 0,
        farra: 0,
        tb_won: 0,
        tb_played: 0,
        hammer_won: 0,
        seq: [],
        best_division_weight: 0,
        division_weight_total: 0,
        division_weight_count: 0
      };
    }

    return playerStats[k];
  }

  function addRowToGroup(r, leagueId, division, sourceDivision) {
    const player = safe(r[idx['player']]);
    if (!leagueId || !division || !player) return;

    const s = ensurePlayer(leagueId, division, player);

    const dw = divisionWeight(sourceDivision || division);
    s.best_division_weight = Math.max(s.best_division_weight, dw);
    s.division_weight_total += dw;
    s.division_weight_count += 1;

    const won = num(r[idx['won_match']]);
    const lost = num(r[idx['lost_match']]);
    const cbs = num(r[idx['cbs']]);
    const farra = num(r[idx['farra']]);
    const chrono = num(r[idx['chronological_index']]);

    s.matches += won + lost;
    s.wins += won;
    s.losses += lost;
    s.cbs += cbs;
    s.farra += farra;

    if (won === 1 || lost === 1) {
      s.seq.push({
        chronological_index: chrono,
        result: won === 1 ? 'W' : 'L'
      });
    }

    const setResults = [
      safe(r[idx['set1_result']]),
      safe(r[idx['set2_result']]),
      safe(r[idx['set3_result']])
    ];

    const setTypes = [
      safe(r[idx['set1_type']]),
      safe(r[idx['set2_type']]),
      safe(r[idx['set3_type']])
    ];

    for (let i = 0; i < 3; i++) {
      if (!setTypes[i]) continue;

      if (setTypes[i] === 'TB') {
        s.tb_played += 1;

        if (setResults[i] === 'W') {
          s.tb_won += 1;
        }
      }

      if (setTypes[i] === 'HAMMER' && setResults[i] === 'W') {
        s.hammer_won += 1;
      }
    }
  }

  rows.forEach(r => {
    const rowLeagueId = safe(r[idx['campeonato']]);
    const rowDivision = safe(r[idx['division']]).toUpperCase();

    if (!rowLeagueId || !rowDivision) return;

    // Temporada específica + división específica
    addRowToGroup(r, rowLeagueId, rowDivision, rowDivision);

    // Temporada específica + todas las divisiones
    addRowToGroup(r, rowLeagueId, 'ALL', rowDivision);

    // Histórico todas las temporadas + división específica
    addRowToGroup(r, 'ALL', rowDivision, rowDivision);

    // Histórico todas las temporadas + todas las divisiones
    addRowToGroup(r, 'ALL', 'ALL', rowDivision);
  });

  function bestWinStreak(seq) {
    if (!seq || seq.length === 0) return 0;

    const ordered = [...seq].sort((a, b) => a.chronological_index - b.chronological_index);

    let best = 0;
    let current = 0;

    ordered.forEach(item => {
      if (item.result === 'W') {
        current++;
        if (current > best) best = current;
      } else {
        current = 0;
      }
    });

    return best;
  }

  function currentWinStreak(seq) {
    if (!seq || seq.length === 0) return 0;

    const ordered = [...seq].sort((a, b) => b.chronological_index - a.chronological_index);

    let current = 0;

    for (const item of ordered) {
      if (item.result === 'W') {
        current++;
      } else {
        break;
      }
    }

    return current;
  }

  function pct(v) {
    return Math.round(v * 1000) / 10;
  }

  function addTopRows(outRows, groupItems, category, valueGetter, labelGetter, extraGetter, options) {
    const opts = options || {};
    const minValue = opts.minValue ?? 0;
    const maxRows = opts.maxRows || 3;

    let items = groupItems
      .map(item => ({
        item,
        value: valueGetter(item)
      }))
      .filter(x => x.value > minValue);

    if (opts.minMatches) {
      const withMin = items.filter(x => x.item.matches >= opts.minMatches);
      if (withMin.length >= 3) {
        items = withMin;
      }
    }

    items.sort((a, b) => {
      if (b.value !== a.value) return b.value - a.value;

      if (b.item.best_division_weight !== a.item.best_division_weight) {
        return b.item.best_division_weight - a.item.best_division_weight;
      }

      const avgDivA = a.item.division_weight_count > 0
        ? a.item.division_weight_total / a.item.division_weight_count
        : 0;

      const avgDivB = b.item.division_weight_count > 0
        ? b.item.division_weight_total / b.item.division_weight_count
        : 0;

      if (avgDivB !== avgDivA) return avgDivB - avgDivA;

      if (b.item.matches !== a.item.matches) return b.item.matches - a.item.matches;

      return a.item.player.localeCompare(b.item.player);
    });

    items.slice(0, maxRows).forEach((x, i) => {
      outRows.push([
        x.item.league_id,
        x.item.division,
        category,
        i + 1,
        x.item.player,
        x.value,
        labelGetter(x.item, x.value),
        extraGetter ? extraGetter(x.item, x.value) : ''
      ]);
    });
  }

  const grouped = {};

  Object.values(playerStats).forEach(s => {
    const gk = groupKey(s.league_id, s.division);
    if (!grouped[gk]) grouped[gk] = [];
    grouped[gk].push(s);
  });

  const outRows = [];

  Object.keys(grouped).sort().forEach(gk => {
    const items = grouped[gk];

    addTopRows(
      outRows,
      items,
      'TOP_WIN_RATE',
      s => s.matches > 0 ? s.wins / s.matches : 0,
      (s, v) => `${pct(v)}%`,
      s => `${s.wins}-${s.losses} · ${s.matches} PJ`,
      { minValue: 0, minMatches: 3 }
    );

    addTopRows(
      outRows,
      items,
      'MOST_WINS',
      s => s.wins,
      s => `${s.wins} victorias`,
      s => `${s.wins}-${s.losses} · ${s.matches} PJ`,
      { minValue: 0 }
    );

    addTopRows(
      outRows,
      items,
      'BEST_STREAK',
      s => bestWinStreak(s.seq),
      (s, v) => `W${v}`,
      s => `${s.wins}-${s.losses} · ${s.matches} PJ`,
      { minValue: 0 }
    );

    addTopRows(
      outRows,
      items,
      'BEST_CURRENT_STREAK',
      s => currentWinStreak(s.seq),
      (s, v) => `W${v}`,
      s => `${s.wins}-${s.losses} · ${s.matches} PJ`,
      { minValue: 0 }
    );

    addTopRows(
      outRows,
      items,
      'MR_COMEBACK',
      s => s.cbs,
      s => `${s.cbs} comeback${s.cbs === 1 ? '' : 's'}`,
      s => `${s.matches} PJ`,
      { minValue: 0 }
    );

    addTopRows(
      outRows,
      items,
      'MR_FARRA',
      s => s.farra,
      s => `${s.farra} farra${s.farra === 1 ? '' : 's'}`,
      s => `${s.matches} PJ`,
      { minValue: 0 }
    );

    addTopRows(
      outRows,
      items,
      'MR_TIE_BREAK',
      s => s.tb_won,
      s => `${s.tb_won} (${s.tb_played})`,
      s => `${s.tb_played} TB jugados`,
      { minValue: 0 }
    );

    addTopRows(
      outRows,
      items,
      'MR_HAMMER',
      s => s.hammer_won,
      s => `${s.hammer_won} hammer ganado${s.hammer_won === 1 ? '' : 's'}`,
      s => `${s.matches} PJ`,
      { minValue: 0 }
    );
  });

  out.clearContents();
  out.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, outputHeaders.length).setValues(outRows);
  }

  notifyUser(
    'league_stats_rankings reconstruida\n' +
    'Filas generadas: ' + outRows.length
  );
}

function buildLeagueStatsSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('matches_raw');
  const standingsSheet = ss.getSheetByName('division_standings');
  const out = ss.getSheetByName('league_stats_summary');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!standingsSheet) throw new Error('No existe la hoja division_standings');
  if (!out) throw new Error('No existe la hoja league_stats_summary');

  const data = raw.getDataRange().getValues();
  if (data.length < 2) throw new Error('matches_raw no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const requiredHeaders = [
    'campeonato', 'division',
    'pareja1_j1', 'pareja1_j2', 'pareja2_j1', 'pareja2_j2',
    'winner',
    'set1_p1', 'set1_p2',
    'set2_p1', 'set2_p2',
    'set3_p1', 'set3_p2',
    'result_type'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en matches_raw: ' + missing.join(', '));
  }

  const outputHeaders = [
    'league_id',
    'division',
    'matches_played',
    'wo_matches',
    'sets_played',
    'games_played',
    'matches_2s',
    'matches_3s',
    'pct_2s',
    'pct_3s',
    'avg_sets_per_match',
    'avg_games_per_match',
    'tie_break_sets',
'hammer_sets',
'cold_sets',
'tight_sets',
'closed_sets',
'closed_sets_pct',
'comeback_matches',
'farra_matches',
    'competitive_score',
    'score_match_balance',
    'score_standings_parity',
    'score_intensity',
    'score_drama',
    'score_commitment',
    'competitive_label'
  ];

  const stats = {};

  function safe(v) {
    return v === null || v === undefined ? '' : String(v).trim();
  }

  function hasValue(v) {
    return v !== '' && v !== null && typeof v !== 'undefined';
  }

  function num(v) {
    if (!hasValue(v)) return 0;
    return Number(v);
  }

  function round2(v) {
    return Math.round(v * 100) / 100;
  }

  function clamp(v, minV, maxV) {
    if (v < minV) return minV;
    if (v > maxV) return maxV;
    return v;
  }

  function key(leagueId, division) {
    return `${leagueId}|${division}`;
  }

  function ensureGroup(leagueId, division) {
    const k = key(leagueId, division);

    if (!stats[k]) {
      stats[k] = {
        league_id: leagueId,
        division: division,
        matches_played: 0,
        wo_matches: 0,
wo_not_played: 0,
wo_injury: 0,
sets_played: 0,
        games_played: 0,
        matches_2s: 0,
        matches_3s: 0,
        tie_break_sets: 0,
        hammer_sets: 0,
        cold_sets: 0,
        tight_sets: 0,
        comeback_matches: 0,
        farra_matches: 0
      };
    }

    return stats[k];
  }

  function addSetType(group, a, b) {
    const hi = Math.max(a, b);
    const lo = Math.min(a, b);

    if (hi === 7 && lo === 6) {
      group.tie_break_sets += 1;
    } else if (hi === 6 && (lo === 0 || lo === 1)) {
      group.hammer_sets += 1;
    } else if (hi === 6 && (lo === 2 || lo === 3)) {
      group.cold_sets += 1;
    } else if ((hi === 6 && lo === 4) || (hi === 7 && lo === 5)) {
      group.tight_sets += 1;
    }
  }

  function processGroup(group, payload) {
    const {
      isWO,
      winner,
      s1a, s1b,
      s2a, s2b,
      s3aRaw, s3bRaw,
      s3a, s3b
    } = payload;

    if (isWO) {
  group.wo_matches += 1;

  if (payload.resultType === 'WO_2_0') {
    group.wo_not_played += 1;
  } else if (
    payload.resultType === 'WO_INJURY_2_0' ||
    payload.resultType === 'WO_INJURY_2_1'
  ) {
    group.wo_injury += 1;
  }

  return;
}

    const hasThird = hasValue(s3aRaw) && hasValue(s3bRaw);

    const setPairs = [
      [s1a, s1b],
      [s2a, s2b]
    ];

    if (hasThird) {
      setPairs.push([s3a, s3b]);
    }

    group.matches_played += 1;
    group.sets_played += setPairs.length;

    if (hasThird) {
      group.matches_3s += 1;
    } else {
      group.matches_2s += 1;
    }

    setPairs.forEach(([a, b]) => {
      group.games_played += a + b;
      addSetType(group, a, b);
    });

    const firstSetWinner = s1a > s1b ? 'P1' : 'P2';

    if (hasThird && winner !== firstSetWinner) {
      group.comeback_matches += 1;
      group.farra_matches += 1;
    }
  }

  function buildParityScores() {
    const standingsData = standingsSheet.getDataRange().getValues();
    if (standingsData.length < 2) return {};

    const sh = standingsData[0].map(h => String(h).trim());
    const sr = standingsData.slice(1);

    const sIdx = {};
    sh.forEach((h, i) => sIdx[h] = i);

    if (!('league_id' in sIdx) || !('division' in sIdx) || !('pts' in sIdx)) {
      throw new Error('division_standings debe tener columnas league_id, division y pts');
    }

    const groups = {};

    sr.forEach(r => {
      const leagueId = safe(r[sIdx['league_id']]);
      const division = safe(r[sIdx['division']]).toUpperCase();
      const pts = Number(r[sIdx['pts']] || 0);

      if (!leagueId || !division) return;

      const k = key(leagueId, division);
      if (!groups[k]) groups[k] = [];
      groups[k].push(pts);
    });

    const parity = {};
    const leagueDivScores = {};

    Object.keys(groups).forEach(k => {
      const ptsArr = groups[k].filter(v => !isNaN(v));

      if (ptsArr.length < 2) {
        parity[k] = 0;
        return;
      }

      ptsArr.sort((a, b) => b - a);

      const maxPts = ptsArr[0];
      const minPts = ptsArr[ptsArr.length - 1];
      const gap = maxPts - minPts;

      let score = 5;

      if (gap <= 3) score = 25;
      else if (gap <= 6) score = 20;
      else if (gap <= 9) score = 15;
      else if (gap <= 12) score = 10;
      else score = 5;

      parity[k] = score;

      const leagueId = k.split('|')[0];
      if (!leagueDivScores[leagueId]) leagueDivScores[leagueId] = [];
      leagueDivScores[leagueId].push(score);
      // Acumula también paridad histórica por división
const parts = k.split('|');
const divId = parts[1];

const histDivKey = key('ALL', divId);
if (!leagueDivScores[histDivKey]) leagueDivScores[histDivKey] = [];
leagueDivScores[histDivKey].push(score);
    });

    Object.keys(leagueDivScores).forEach(groupId => {
  const arr = leagueDivScores[groupId];
  const avg = arr.reduce((a, b) => a + b, 0) / arr.length;

  // Caso normal: leagueId real → leagueId | ALL
  if (!groupId.includes('|')) {
    parity[key(groupId, 'ALL')] = round2(avg);
  }

  // Caso histórico por división: ALL|A, ALL|B, etc.
  if (groupId.includes('|')) {
    parity[groupId] = round2(avg);
  }
});

// Paridad histórica total: promedio de todas las divisiones históricas
const histAllKeys = ['A', 'B', 'C', 'D']
  .map(d => key('ALL', d))
  .filter(k => parity[k] !== undefined);

if (histAllKeys.length > 0) {
  const avgHistAll = histAllKeys.reduce((acc, k) => acc + Number(parity[k] || 0), 0) / histAllKeys.length;
  parity[key('ALL', 'ALL')] = round2(avgHistAll);
}

    return parity;
  }

  function competitiveLabel(score) {
    if (score >= 90) return 'Histórica';
    if (score >= 75) return 'Muy competitiva';
    if (score >= 60) return 'Competitiva';
    if (score >= 40) return 'Irregular';
    return 'Dominada';
  }

  function computeScore(s, parityScore) {
  const realMatches = s.matches_played;
  const totalClosed = s.matches_played + s.wo_matches;

  if (realMatches <= 0 && totalClosed <= 0) {
    return {
      total: 0,
      matchBalance: 0,   // Largo de partidos
      parity: 0,         // Paridad de tabla
      intensity: 0,      // Apretura de sets
      drama: 0,
      commitment: 0,
      label: 'Sin datos'
    };
  }

  const pct3 = realMatches > 0 ? s.matches_3s / realMatches : 0;

  const closedSetRate = s.sets_played > 0
    ? (s.tight_sets + s.tie_break_sets) / s.sets_played
    : 0;

  const tbRate = s.sets_played > 0
    ? s.tie_break_sets / s.sets_played
    : 0;

  const comebackRate = realMatches > 0
    ? s.comeback_matches / realMatches
    : 0;

  const woNotPlayedRate = totalClosed > 0
    ? (s.wo_not_played || 0) / totalClosed
    : 0;

  const woInjuryRate = totalClosed > 0
    ? (s.wo_injury || 0) / totalClosed
    : 0;

  // 1. Largo de partidos — máximo 45 pts
  // 35% de partidos a 3 sets equivale al máximo.
  const longMatchScore = clamp((pct3 / 0.45) * 45, 0, 45);

  // 2. Apretura de sets — máximo 40 pts
  // La base es % de sets cerrados. El tie break suma intensidad extra, pero no domina.
  const setTightnessScore =
  clamp((closedSetRate / 0.70) * 34, 0, 34) +
  clamp((tbRate / 0.25) * 6, 0, 6);

  // 3. Paridad de tabla — máximo 7 pts
  // buildParityScores entrega hasta 25 pts. Lo normalizamos a 7.
  const parity = clamp(((parityScore || 0) / 25) * 7, 0, 7);

  // 4. Drama competitivo — máximo 5 pts
  const drama = clamp((comebackRate / 0.25) * 5, 0, 5);

  // 5. Compromiso — máximo 3 pts
  const commitmentPenalty =
    clamp((woNotPlayedRate / 0.25) * 2.5, 0, 2.5) +
    clamp((woInjuryRate / 0.25) * 0.5, 0, 0.5);

  const commitment = clamp(3 - commitmentPenalty, 0, 3);

  const total = Math.round(
    longMatchScore +
    setTightnessScore +
    parity +
    drama +
    commitment
  );

  return {
    total,
    matchBalance: round2(longMatchScore),
    parity: round2(parity),
    intensity: round2(setTightnessScore),
    drama: round2(drama),
    commitment: round2(commitment),
    label: competitiveLabel(total)
  };
}

  rows.forEach(r => {
    const leagueId = safe(r[idx['campeonato']]);
    const division = safe(r[idx['division']]).toUpperCase();

    const p1a = safe(r[idx['pareja1_j1']]);
    const p1b = safe(r[idx['pareja1_j2']]);
    const p2a = safe(r[idx['pareja2_j1']]);
    const p2b = safe(r[idx['pareja2_j2']]);

    const winner = safe(r[idx['winner']]).toUpperCase();
    const resultType = safe(r[idx['result_type']]).toUpperCase();

    const isWO =
      resultType === 'WO_2_0' ||
      resultType === 'WO_INJURY_2_0' ||
      resultType === 'WO_INJURY_2_1';

    const s1aRaw = r[idx['set1_p1']];
    const s1bRaw = r[idx['set1_p2']];
    const s2aRaw = r[idx['set2_p1']];
    const s2bRaw = r[idx['set2_p2']];
    const s3aRaw = r[idx['set3_p1']];
    const s3bRaw = r[idx['set3_p2']];

    const hasPlayers =
      leagueId !== '' &&
      division !== '' &&
      p1a !== '' && p1b !== '' && p2a !== '' && p2b !== '';

    const hasWinner = winner === 'P1' || winner === 'P2';

    const hasRealScore =
      hasValue(s1aRaw) &&
      hasValue(s1bRaw) &&
      hasValue(s2aRaw) &&
      hasValue(s2bRaw);

    const isClosedForSummary = hasPlayers && hasWinner && (hasRealScore || isWO);

    if (!isClosedForSummary) return;

    const payload = {
  isWO,
  resultType,
  winner,
  s1a: num(s1aRaw),
      s1b: num(s1bRaw),
      s2a: num(s2aRaw),
      s2b: num(s2bRaw),
      s3aRaw,
      s3bRaw,
      s3a: num(s3aRaw),
      s3b: num(s3bRaw)
    };

    // Temporada específica + división específica
processGroup(ensureGroup(leagueId, division), payload);

// Temporada específica + todas las divisiones
processGroup(ensureGroup(leagueId, 'ALL'), payload);

// Histórico todas las temporadas + división específica
processGroup(ensureGroup('ALL', division), payload);

// Histórico todas las temporadas + todas las divisiones
processGroup(ensureGroup('ALL', 'ALL'), payload);
  });

  Logger.log('STATS KEYS: ' + Object.keys(stats).sort().join(' | '));

  const parityScores = buildParityScores();

  const outRows = Object.values(stats)
    .sort((a, b) => {
      if (a.league_id < b.league_id) return -1;
      if (a.league_id > b.league_id) return 1;

      const order = { ALL: 0, A: 1, B: 2, C: 3, D: 4 };
      return (order[a.division] || 99) - (order[b.division] || 99);
    })
    .map(s => {
      const totalReal = s.matches_played;

      const pct2 = totalReal > 0 ? s.matches_2s / totalReal : '';
      const pct3 = totalReal > 0 ? s.matches_3s / totalReal : '';
      const avgSets = totalReal > 0 ? s.sets_played / totalReal : '';
      const avgGames = totalReal > 0 ? s.games_played / totalReal : '';

      const score = computeScore(s, parityScores[key(s.league_id, s.division)]);

      return [
        s.league_id,
        s.division,
        s.matches_played,
        s.wo_matches,
        s.sets_played,
        s.games_played,
        s.matches_2s,
        s.matches_3s,
        pct2 === '' ? '' : pct2,
        pct3 === '' ? '' : pct3,
        avgSets === '' ? '' : round2(avgSets),
        avgGames === '' ? '' : round2(avgGames),
        s.tie_break_sets,
s.hammer_sets,
s.cold_sets,
s.tight_sets,
s.tight_sets + s.tie_break_sets,
s.sets_played > 0 ? round2((s.tight_sets + s.tie_break_sets) / s.sets_played) : '',
s.comeback_matches,
s.farra_matches,
score.total,
        score.matchBalance,
        score.parity,
        score.intensity,
        score.drama,
        score.commitment,
        score.label
      ];
    });

  out.clearContents();
  out.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, outputHeaders.length).setValues(outRows);

    out.getRange(2, 9, outRows.length, 2).setNumberFormat('0.0%');   // pct_2s, pct_3s
out.getRange(2, 11, outRows.length, 2).setNumberFormat('0.00');  // avg_sets, avg_games
out.getRange(2, 18, outRows.length, 1).setNumberFormat('0.0%');  // closed_sets_pct
out.getRange(2, 21, outRows.length, 6).setNumberFormat('0.00');  // score components
  }

  notifyUser(
    'league_stats_summary reconstruida\n' +
    'Filas generadas: ' + outRows.length
  );
}
function buildStatsRachas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('matches_raw');
  const out = ss.getSheetByName('stats_rachas');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!out) throw new Error('No existe la hoja stats_rachas');

  const data = raw.getDataRange().getValues();
  if (data.length < 2) throw new Error('matches_raw no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  const orderedRows = sortRowsByPlayedOrder(rows, idx);

  const requiredHeaders = [
    'division',
    'pareja1_j1', 'pareja1_j2', 'pareja2_j1', 'pareja2_j2',
    'winner', 'set1_p1', 'set1_p2', 'set2_p1', 'set2_p2'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en matches_raw: ' + missing.join(', '));
  }

  const outputHeaders = [
    'player', 'pj_reales',
    'seq_5', 'g_5', 'p_5', 'wr_5',
    'seq_10', 'g_10', 'p_10', 'wr_10',
    'racha_actual', 'tipo_racha_actual', 'largo_racha_actual',
    'mejor_racha_g', 'mejor_racha_p',
    'seq_hist', 'g_hist', 'p_hist', 'wr_hist'
  ];

  const playerSeq = {};

  for (let i = 0; i < orderedRows.length; i++) {
  const r = orderedRows[i];

    const division = safe(r[idx['division']]);
    const p1a = safe(r[idx['pareja1_j1']]);
    const p1b = safe(r[idx['pareja1_j2']]);
    const p2a = safe(r[idx['pareja2_j1']]);
    const p2b = safe(r[idx['pareja2_j2']]);
    const winner = safe(r[idx['winner']]);

    const s1p1 = r[idx['set1_p1']];
    const s1p2 = r[idx['set1_p2']];
    const s2p1 = r[idx['set2_p1']];
    const s2p2 = r[idx['set2_p2']];

    const played =
      division !== '' &&
      p1a !== '' && p1b !== '' && p2a !== '' && p2b !== '' &&
      (winner === 'P1' || winner === 'P2') &&
      s1p1 !== '' && s1p1 !== null &&
      s1p2 !== '' && s1p2 !== null &&
      s2p1 !== '' && s2p1 !== null &&
      s2p2 !== '' && s2p2 !== null;

    if (!played) continue;

    addSeqResult(playerSeq, p1a, winner === 'P1' ? 'W' : 'L');
    addSeqResult(playerSeq, p1b, winner === 'P1' ? 'W' : 'L');
    addSeqResult(playerSeq, p2a, winner === 'P2' ? 'W' : 'L');
    addSeqResult(playerSeq, p2b, winner === 'P2' ? 'W' : 'L');
  }

  const players = Object.keys(playerSeq).sort();
  const outRows = [];

  players.forEach(player => {
    const seqArr = playerSeq[player];
    const pj = seqArr.length;

    const last5Arr = seqArr.slice(-5);
    const last10Arr = seqArr.slice(-10);

    const g5 = countValue(last5Arr, 'W');
    const p5 = countValue(last5Arr, 'L');
    const wr5 = last5Arr.length ? g5 / last5Arr.length : '';

    const g10 = countValue(last10Arr, 'W');
    const p10 = countValue(last10Arr, 'L');
    const wr10 = last10Arr.length ? g10 / last10Arr.length : '';

    const gHist = countValue(seqArr, 'W');
    const pHist = countValue(seqArr, 'L');
    const wrHist = pj ? gHist / pj : '';

    const streak = currentStreak(seqArr);
    const best = bestStreaks(seqArr);

    outRows.push([
      player,
      pj,
      seqToDisplay(last5Arr),
      g5,
      p5,
      wr5,
      seqToDisplay(last10Arr),
      g10,
      p10,
      wr10,
      streak.label,
      streak.type,
      streak.length,
      best.bestW,
      best.bestL,
      seqToDisplay(seqArr),
      gHist,
      pHist,
      wrHist
    ]);
  });

  out.clearContents();
  out.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, outputHeaders.length).setValues(outRows);
    out.getRange(2, 6, outRows.length, 1).setNumberFormat('0.0%');
    out.getRange(2, 10, outRows.length, 1).setNumberFormat('0.0%');
    out.getRange(2, 19, outRows.length, 1).setNumberFormat('0.0%');
  }

  notifyUser(
    'stats_rachas reconstruida\n' +
    'Jugadores procesados: ' + outRows.length
  );
}

function addSeqResult(playerSeq, player, resultCode) {
  if (!player) return;
  if (!playerSeq[player]) playerSeq[player] = [];
  playerSeq[player].push(resultCode);
}

function countValue(arr, val) {
  let n = 0;
  for (let i = 0; i < arr.length; i++) {
    if (arr[i] === val) n++;
  }
  return n;
}

function seqToDisplay(arr) {
  return arr.join('-');
}

function currentStreak(seqArr) {
  if (!seqArr.length) {
    return { label: '', type: '', length: 0 };
  }

  const last = seqArr[seqArr.length - 1];
  let len = 0;

  for (let i = seqArr.length - 1; i >= 0; i--) {
    if (seqArr[i] === last) len++;
    else break;
  }

  return {
    label: last + String(len),
    type: last,
    length: len
  };
}

function bestStreaks(seqArr) {
  let bestW = 0;
  let bestL = 0;
  let curW = 0;
  let curL = 0;

  for (let i = 0; i < seqArr.length; i++) {
    const v = seqArr[i];

    if (v === 'W') {
      curW++;
      curL = 0;
    } else if (v === 'L') {
      curL++;
      curW = 0;
    }

    if (curW > bestW) bestW = curW;
    if (curL > bestL) bestL = curL;
  }

  return { bestW, bestL };
}
function buildStatsIndivBlock1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('player_match_log');
  const out = ss.getSheetByName('stats_indiv');

  if (!log) throw new Error('No existe la hoja player_match_log');
  if (!out) throw new Error('No existe la hoja stats_indiv');

  const data = log.getDataRange().getValues();
  if (data.length < 2) throw new Error('player_match_log no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const requiredHeaders = [
    'player','division',
    'won_match','lost_match',
    'sets_won','sets_lost',
    'games_won','games_lost'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en player_match_log: ' + missing.join(', '));
  }

  const outputHeaders = [
    'player',
    'match_g','match_p','match_wr',
    'sets_g','sets_p','sets_wr',
    'games_g','games_p','games_wr',
    'a_match_g','a_match_p','a_match_wr',
    'a_sets_g','a_sets_p','a_sets_wr',
    'a_games_g','a_games_p','a_games_wr',
    'b_match_g','b_match_p','b_match_wr',
    'b_sets_g','b_sets_p','b_sets_wr',
    'b_games_g','b_games_p','b_games_wr',
    'c_match_g','c_match_p','c_match_wr',
    'c_sets_g','c_sets_p','c_sets_wr',
    'c_games_g','c_games_p','c_games_wr',
    'd_match_g','d_match_p','d_match_wr',
    'd_sets_g','d_sets_p','d_sets_wr',
    'd_games_g','d_games_p','d_games_wr'
  ];

  const stats = {};

  rows.forEach(r => {
    const player = safe(r[idx['player']]);
    const div = safe(r[idx['division']]).toUpperCase();

    if (!player) return;

    if (!stats[player]) {
      stats[player] = emptyStatsIndivBlock1(player);
    }

    const s = stats[player];

    const wonMatch = Number(r[idx['won_match']] || 0);
    const lostMatch = Number(r[idx['lost_match']] || 0);
    const setsWon = Number(r[idx['sets_won']] || 0);
    const setsLost = Number(r[idx['sets_lost']] || 0);
    const gamesWon = Number(r[idx['games_won']] || 0);
    const gamesLost = Number(r[idx['games_lost']] || 0);

    s.match_g += wonMatch;
    s.match_p += lostMatch;
    s.sets_g += setsWon;
    s.sets_p += setsLost;
    s.games_g += gamesWon;
    s.games_p += gamesLost;

    if (div === 'A') {
      s.a_match_g += wonMatch;
      s.a_match_p += lostMatch;
      s.a_sets_g += setsWon;
      s.a_sets_p += setsLost;
      s.a_games_g += gamesWon;
      s.a_games_p += gamesLost;
    } else if (div === 'B') {
      s.b_match_g += wonMatch;
      s.b_match_p += lostMatch;
      s.b_sets_g += setsWon;
      s.b_sets_p += setsLost;
      s.b_games_g += gamesWon;
      s.b_games_p += gamesLost;
    } else if (div === 'C') {
      s.c_match_g += wonMatch;
      s.c_match_p += lostMatch;
      s.c_sets_g += setsWon;
      s.c_sets_p += setsLost;
      s.c_games_g += gamesWon;
      s.c_games_p += gamesLost;
    } else if (div === 'D') {
      s.d_match_g += wonMatch;
      s.d_match_p += lostMatch;
      s.d_sets_g += setsWon;
      s.d_sets_p += setsLost;
      s.d_games_g += gamesWon;
      s.d_games_p += gamesLost;
    }
  });

  const players = Object.keys(stats).sort();
  const outRows = players.map(player => {
    const s = stats[player];

    s.match_wr = rate(s.match_g, s.match_g + s.match_p);
    s.sets_wr = rate(s.sets_g, s.sets_g + s.sets_p);
    s.games_wr = rate(s.games_g, s.games_g + s.games_p);

    s.a_match_wr = rate(s.a_match_g, s.a_match_g + s.a_match_p);
    s.a_sets_wr = rate(s.a_sets_g, s.a_sets_g + s.a_sets_p);
    s.a_games_wr = rate(s.a_games_g, s.a_games_g + s.a_games_p);

    s.b_match_wr = rate(s.b_match_g, s.b_match_g + s.b_match_p);
    s.b_sets_wr = rate(s.b_sets_g, s.b_sets_g + s.b_sets_p);
    s.b_games_wr = rate(s.b_games_g, s.b_games_g + s.b_games_p);

    s.c_match_wr = rate(s.c_match_g, s.c_match_g + s.c_match_p);
    s.c_sets_wr = rate(s.c_sets_g, s.c_sets_g + s.c_sets_p);
    s.c_games_wr = rate(s.c_games_g, s.c_games_g + s.c_games_p);

    s.d_match_wr = rate(s.d_match_g, s.d_match_g + s.d_match_p);
    s.d_sets_wr = rate(s.d_sets_g, s.d_sets_g + s.d_sets_p);
    s.d_games_wr = rate(s.d_games_g, s.d_games_g + s.d_games_p);

    return [
      s.player,
      s.match_g, s.match_p, s.match_wr,
      s.sets_g, s.sets_p, s.sets_wr,
      s.games_g, s.games_p, s.games_wr,
      s.a_match_g, s.a_match_p, s.a_match_wr,
      s.a_sets_g, s.a_sets_p, s.a_sets_wr,
      s.a_games_g, s.a_games_p, s.a_games_wr,
      s.b_match_g, s.b_match_p, s.b_match_wr,
      s.b_sets_g, s.b_sets_p, s.b_sets_wr,
      s.b_games_g, s.b_games_p, s.b_games_wr,
      s.c_match_g, s.c_match_p, s.c_match_wr,
      s.c_sets_g, s.c_sets_p, s.c_sets_wr,
      s.c_games_g, s.c_games_p, s.c_games_wr,
      s.d_match_g, s.d_match_p, s.d_match_wr,
      s.d_sets_g, s.d_sets_p, s.d_sets_wr,
      s.d_games_g, s.d_games_p, s.d_games_wr
    ];
  });

  out.clearContents();
  out.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, outputHeaders.length).setValues(outRows);

    const pctCols = [4,7,10,13,16,19,22,25,28,31,34,37,40,43,46];
    pctCols.forEach(col => {
      out.getRange(2, col, outRows.length, 1).setNumberFormat('0.0%');
    });
  }

  notifyUser(
    'stats_indiv Bloque 1 reconstruido\n' +
    'Jugadores procesados: ' + outRows.length
  );
}

function emptyStatsIndivBlock1(player) {
  return {
    player: player,

    match_g: 0, match_p: 0, match_wr: '',
    sets_g: 0, sets_p: 0, sets_wr: '',
    games_g: 0, games_p: 0, games_wr: '',

    a_match_g: 0, a_match_p: 0, a_match_wr: '',
    a_sets_g: 0, a_sets_p: 0, a_sets_wr: '',
    a_games_g: 0, a_games_p: 0, a_games_wr: '',

    b_match_g: 0, b_match_p: 0, b_match_wr: '',
    b_sets_g: 0, b_sets_p: 0, b_sets_wr: '',
    b_games_g: 0, b_games_p: 0, b_games_wr: '',

    c_match_g: 0, c_match_p: 0, c_match_wr: '',
    c_sets_g: 0, c_sets_p: 0, c_sets_wr: '',
    c_games_g: 0, c_games_p: 0, c_games_wr: '',

    d_match_g: 0, d_match_p: 0, d_match_wr: '',
    d_sets_g: 0, d_sets_p: 0, d_sets_wr: '',
    d_games_g: 0, d_games_p: 0, d_games_wr: ''
  };
}

function rate(num, den) {
  return den > 0 ? num / den : '';
}
function buildStatsIndivBlock2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('player_match_log');
  const out = ss.getSheetByName('stats_indiv');

  if (!log) throw new Error('No existe la hoja player_match_log');
  if (!out) throw new Error('No existe la hoja stats_indiv');

  const logData = log.getDataRange().getValues();
  if (logData.length < 2) throw new Error('player_match_log no tiene datos');

  const logHeaders = logData[0].map(h => String(h).trim());
  const logRows = logData.slice(1);

  const logIdx = {};
  logHeaders.forEach((h, i) => logIdx[h] = i);

  const requiredLogHeaders = [
    'player', 'match_format', 'won_2s', 'lost_2s', 'won_3s', 'lost_3s', 'cbs', 'farra'
  ];

  const missing = requiredLogHeaders.filter(h => !(h in logIdx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en player_match_log: ' + missing.join(', '));
  }

  const outData = out.getDataRange().getValues();
  if (outData.length < 2) throw new Error('stats_indiv no tiene Bloque 1 cargado');

  const outHeaders = outData[0].map(h => String(h).trim());
  const outRows = outData.slice(1);

  const outIdx = {};
  outHeaders.forEach((h, i) => outIdx[h] = i);

  if (!('player' in outIdx)) {
    throw new Error('stats_indiv no tiene la columna player');
  }

  const block2Headers = [
    'tot_2s', 'pj_2s_pct', 'tot_3s', 'pj_3s_pct',
    'pg_2s', 'pp_2s', 'wr_2s',
    'pg_3s', 'pp_3s', 'wr_3s',
    'cbs_total', 'farra_total'
  ];

  // Si no existen, agregarlas al final
  let currentLastCol = outHeaders.length;
  block2Headers.forEach(h => {
    if (!(h in outIdx)) {
      currentLastCol++;
      out.getRange(1, currentLastCol).setValue(h);
      outIdx[h] = currentLastCol - 1;
    }
  });

  const stats = {};

  logRows.forEach(r => {
    const player = safe(r[logIdx['player']]);
    if (!player) return;

    if (!stats[player]) {
      stats[player] = {
        tot_2s: 0,
        tot_3s: 0,
        pg_2s: 0,
        pp_2s: 0,
        pg_3s: 0,
        pp_3s: 0,
        cbs_total: 0,
        farra_total: 0
      };
    }

    const s = stats[player];

    const won2s = Number(r[logIdx['won_2s']] || 0);
    const lost2s = Number(r[logIdx['lost_2s']] || 0);
    const won3s = Number(r[logIdx['won_3s']] || 0);
    const lost3s = Number(r[logIdx['lost_3s']] || 0);
    const cbs = Number(r[logIdx['cbs']] || 0);
    const farra = Number(r[logIdx['farra']] || 0);

    s.pg_2s += won2s;
    s.pp_2s += lost2s;
    s.pg_3s += won3s;
    s.pp_3s += lost3s;
    s.cbs_total += cbs;
    s.farra_total += farra;
  });

  const playerRowMap = {};
  outRows.forEach((r, i) => {
    const player = safe(r[outIdx['player']]);
    if (player) playerRowMap[player] = i + 2; // +2 because sheet rows start at 2
  });

  Object.keys(playerRowMap).forEach(player => {
    const rowNum = playerRowMap[player];
    const s = stats[player] || {
      tot_2s: 0,
      tot_3s: 0,
      pg_2s: 0,
      pp_2s: 0,
      pg_3s: 0,
      pp_3s: 0,
      cbs_total: 0,
      farra_total: 0
    };

    s.tot_2s = s.pg_2s + s.pp_2s;
    s.tot_3s = s.pg_3s + s.pp_3s;

    const totalMatches = s.tot_2s + s.tot_3s;

    const pj_2s_pct = totalMatches > 0 ? s.tot_2s / totalMatches : '';
    const pj_3s_pct = totalMatches > 0 ? s.tot_3s / totalMatches : '';
    const wr_2s = s.tot_2s > 0 ? s.pg_2s / s.tot_2s : '';
    const wr_3s = s.tot_3s > 0 ? s.pg_3s / s.tot_3s : '';

    const values = [[
      s.tot_2s,
      pj_2s_pct,
      s.tot_3s,
      pj_3s_pct,
      s.pg_2s,
      s.pp_2s,
      wr_2s,
      s.pg_3s,
      s.pp_3s,
      wr_3s,
      s.cbs_total,
      s.farra_total
    ]];

    out.getRange(rowNum, outIdx['tot_2s'] + 1, 1, block2Headers.length).setValues(values);
  });
  // Formatos número entero
  ['tot_2s', 'tot_3s', 'pg_2s', 'pp_2s', 'pg_3s', 'pp_3s', 'cbs_total', 'farra_total'].forEach(h => {
    const col = outIdx[h] + 1;
    out.getRange(2, col, out.getLastRow() - 1, 1).setNumberFormat('0');
  });
  // Formatos %
  ['pj_2s_pct', 'pj_3s_pct', 'wr_2s', 'wr_3s'].forEach(h => {
    const col = outIdx[h] + 1;
    out.getRange(2, col, out.getLastRow() - 1, 1).setNumberFormat('0.0%');
  });

  notifyUser(
    'stats_indiv Bloque 2 reconstruido\n' +
    'Columnas agregadas/actualizadas: ' + block2Headers.length
  );
}
function buildStatsIndivBlock3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('player_match_log');
  const out = ss.getSheetByName('stats_indiv');

  if (!log) throw new Error('No existe la hoja player_match_log');
  if (!out) throw new Error('No existe la hoja stats_indiv');

  const logData = log.getDataRange().getValues();
  if (logData.length < 2) throw new Error('player_match_log no tiene datos');

  const logHeaders = logData[0].map(h => String(h).trim());
  const logRows = logData.slice(1);

  const logIdx = {};
  logHeaders.forEach((h, i) => logIdx[h] = i);

  const requiredLogHeaders = [
    'player', 'set1_result', 'set2_result', 'set3_result', 'sets_won', 'sets_lost'
  ];

  const missing = requiredLogHeaders.filter(h => !(h in logIdx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en player_match_log: ' + missing.join(', '));
  }

  const outData = out.getDataRange().getValues();
  if (outData.length < 2) throw new Error('stats_indiv no tiene datos previos');

  const outHeaders = outData[0].map(h => String(h).trim());
  const outRows = outData.slice(1);

  const outIdx = {};
  outHeaders.forEach((h, i) => outIdx[h] = i);

  if (!('player' in outIdx)) {
    throw new Error('stats_indiv no tiene la columna player');
  }

  const block3Headers = [
    's1_g', 's1_p', 's1_wr',
    's2_g', 's2_p', 's2_wr',
    's3_g', 's3_p', 's3_wr',
    'total_sets_g', 'total_sets_p'
  ];

  let currentLastCol = outHeaders.length;
  block3Headers.forEach(h => {
    if (!(h in outIdx)) {
      currentLastCol++;
      out.getRange(1, currentLastCol).setValue(h);
      outIdx[h] = currentLastCol - 1;
    }
  });

  const stats = {};

  logRows.forEach(r => {
    const player = safe(r[logIdx['player']]);
    if (!player) return;

    if (!stats[player]) {
      stats[player] = {
        s1_g: 0, s1_p: 0,
        s2_g: 0, s2_p: 0,
        s3_g: 0, s3_p: 0,
        total_sets_g: 0, total_sets_p: 0
      };
    }

    const s = stats[player];

    const set1 = safe(r[logIdx['set1_result']]);
    const set2 = safe(r[logIdx['set2_result']]);
    const set3 = safe(r[logIdx['set3_result']]);

    const setsWon = Number(r[logIdx['sets_won']] || 0);
    const setsLost = Number(r[logIdx['sets_lost']] || 0);

    if (set1 === 'W') s.s1_g++;
    else if (set1 === 'L') s.s1_p++;

    if (set2 === 'W') s.s2_g++;
    else if (set2 === 'L') s.s2_p++;

    if (set3 === 'W') s.s3_g++;
    else if (set3 === 'L') s.s3_p++;

    s.total_sets_g += setsWon;
    s.total_sets_p += setsLost;
  });

  const playerRowMap = {};
  outRows.forEach((r, i) => {
    const player = safe(r[outIdx['player']]);
    if (player) playerRowMap[player] = i + 2;
  });

  Object.keys(playerRowMap).forEach(player => {
    const rowNum = playerRowMap[player];
    const s = stats[player] || {
      s1_g: 0, s1_p: 0,
      s2_g: 0, s2_p: 0,
      s3_g: 0, s3_p: 0,
      total_sets_g: 0, total_sets_p: 0
    };

    const s1_wr = rate(s.s1_g, s.s1_g + s.s1_p);
    const s2_wr = rate(s.s2_g, s.s2_g + s.s2_p);
    const s3_wr = rate(s.s3_g, s.s3_g + s.s3_p);

    const values = [[
      s.s1_g, s.s1_p, s1_wr,
      s.s2_g, s.s2_p, s2_wr,
      s.s3_g, s.s3_p, s3_wr,
      s.total_sets_g, s.total_sets_p
    ]];

    out.getRange(rowNum, outIdx['s1_g'] + 1, 1, block3Headers.length).setValues(values);
  });

  ['s1_g', 's1_p', 's2_g', 's2_p', 's3_g', 's3_p', 'total_sets_g', 'total_sets_p'].forEach(h => {
    const col = outIdx[h] + 1;
    out.getRange(2, col, out.getLastRow() - 1, 1).setNumberFormat('0');
  });

  ['s1_wr', 's2_wr', 's3_wr'].forEach(h => {
    const col = outIdx[h] + 1;
    out.getRange(2, col, out.getLastRow() - 1, 1).setNumberFormat('0.0%');
  });

  notifyUser(
    'stats_indiv Bloque 3 reconstruido\n' +
    'Columnas agregadas/actualizadas: ' + block3Headers.length
  );
}
function buildStatsIndivBlock4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('player_match_log');
  const out = ss.getSheetByName('stats_indiv');

  if (!log) throw new Error('No existe la hoja player_match_log');
  if (!out) throw new Error('No existe la hoja stats_indiv');

  const logData = log.getDataRange().getValues();
  if (logData.length < 2) throw new Error('player_match_log no tiene datos');

  const logHeaders = logData[0].map(h => String(h).trim());
  const logRows = logData.slice(1);

  const logIdx = {};
  logHeaders.forEach((h, i) => logIdx[h] = i);

  const requiredLogHeaders = [
    'player',
    'set1_result', 'set2_result', 'set3_result',
    'set1_type', 'set2_type', 'set3_type'
  ];

  const missing = requiredLogHeaders.filter(h => !(h in logIdx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en player_match_log: ' + missing.join(', '));
  }

  const outData = out.getDataRange().getValues();
  if (outData.length < 2) throw new Error('stats_indiv no tiene datos previos');

  const outHeaders = outData[0].map(h => String(h).trim());
  const outRows = outData.slice(1);

  const outIdx = {};
  outHeaders.forEach((h, i) => outIdx[h] = i);

  if (!('player' in outIdx)) {
    throw new Error('stats_indiv no tiene la columna player');
  }

  const block4Headers = [
    'tb_s1_g','tb_s1_p','tb_s1_wr','tb_s2_g','tb_s2_p','tb_s2_wr','tb_s3_g','tb_s3_p','tb_s3_wr',
    'tight_s1_g','tight_s1_p','tight_s1_wr','tight_s2_g','tight_s2_p','tight_s2_wr','tight_s3_g','tight_s3_p','tight_s3_wr',
    'cold_s1_g','cold_s1_p','cold_s1_wr','cold_s2_g','cold_s2_p','cold_s2_wr','cold_s3_g','cold_s3_p','cold_s3_wr',
    'hammer_s1_g','hammer_s1_p','hammer_s1_wr','hammer_s2_g','hammer_s2_p','hammer_s2_wr','hammer_s3_g','hammer_s3_p','hammer_s3_wr',
    'tb_total_g','tb_total_p','tb_total_wr',
    'tight_total_g','tight_total_p','tight_total_wr',
    'cold_total_g','cold_total_p','cold_total_wr',
    'hammer_total_g','hammer_total_p','hammer_total_wr'
  ];

  let currentLastCol = outHeaders.length;
  block4Headers.forEach(h => {
    if (!(h in outIdx)) {
      currentLastCol++;
      out.getRange(1, currentLastCol).setValue(h);
      outIdx[h] = currentLastCol - 1;
    }
  });

  const stats = {};

  logRows.forEach(r => {
    const player = safe(r[logIdx['player']]);
    if (!player) return;

    if (!stats[player]) {
      stats[player] = emptyStatsIndivBlock4();
    }

    const s = stats[player];

    processTypedSet(s, safe(r[logIdx['set1_type']]), safe(r[logIdx['set1_result']]), 's1');
    processTypedSet(s, safe(r[logIdx['set2_type']]), safe(r[logIdx['set2_result']]), 's2');
    processTypedSet(s, safe(r[logIdx['set3_type']]), safe(r[logIdx['set3_result']]), 's3');
  });

  const playerRowMap = {};
  outRows.forEach((r, i) => {
    const player = safe(r[outIdx['player']]);
    if (player) playerRowMap[player] = i + 2;
  });

  Object.keys(playerRowMap).forEach(player => {
    const rowNum = playerRowMap[player];
    const s = stats[player] || emptyStatsIndivBlock4();

    computeTypedRates(s);

    const values = [[
      s.tb_s1_g, s.tb_s1_p, s.tb_s1_wr,
      s.tb_s2_g, s.tb_s2_p, s.tb_s2_wr,
      s.tb_s3_g, s.tb_s3_p, s.tb_s3_wr,

      s.tight_s1_g, s.tight_s1_p, s.tight_s1_wr,
      s.tight_s2_g, s.tight_s2_p, s.tight_s2_wr,
      s.tight_s3_g, s.tight_s3_p, s.tight_s3_wr,

      s.cold_s1_g, s.cold_s1_p, s.cold_s1_wr,
      s.cold_s2_g, s.cold_s2_p, s.cold_s2_wr,
      s.cold_s3_g, s.cold_s3_p, s.cold_s3_wr,

      s.hammer_s1_g, s.hammer_s1_p, s.hammer_s1_wr,
      s.hammer_s2_g, s.hammer_s2_p, s.hammer_s2_wr,
      s.hammer_s3_g, s.hammer_s3_p, s.hammer_s3_wr,

      s.tb_total_g, s.tb_total_p, s.tb_total_wr,
      s.tight_total_g, s.tight_total_p, s.tight_total_wr,
      s.cold_total_g, s.cold_total_p, s.cold_total_wr,
      s.hammer_total_g, s.hammer_total_p, s.hammer_total_wr
    ]];

    out.getRange(rowNum, outIdx['tb_s1_g'] + 1, 1, block4Headers.length).setValues(values);
  });

  // enteros
  block4Headers.filter(h => !h.endsWith('_wr')).forEach(h => {
    const col = outIdx[h] + 1;
    out.getRange(2, col, out.getLastRow() - 1, 1).setNumberFormat('0');
  });

  // porcentajes
  block4Headers.filter(h => h.endsWith('_wr')).forEach(h => {
    const col = outIdx[h] + 1;
    out.getRange(2, col, out.getLastRow() - 1, 1).setNumberFormat('0.0%');
  });

  notifyUser(
    'stats_indiv Bloque 4 reconstruido\n' +
    'Columnas agregadas/actualizadas: ' + block4Headers.length
  );
}

function emptyStatsIndivBlock4() {
  return {
    tb_s1_g: 0, tb_s1_p: 0, tb_s1_wr: '',
    tb_s2_g: 0, tb_s2_p: 0, tb_s2_wr: '',
    tb_s3_g: 0, tb_s3_p: 0, tb_s3_wr: '',

    tight_s1_g: 0, tight_s1_p: 0, tight_s1_wr: '',
    tight_s2_g: 0, tight_s2_p: 0, tight_s2_wr: '',
    tight_s3_g: 0, tight_s3_p: 0, tight_s3_wr: '',

    cold_s1_g: 0, cold_s1_p: 0, cold_s1_wr: '',
    cold_s2_g: 0, cold_s2_p: 0, cold_s2_wr: '',
    cold_s3_g: 0, cold_s3_p: 0, cold_s3_wr: '',

    hammer_s1_g: 0, hammer_s1_p: 0, hammer_s1_wr: '',
    hammer_s2_g: 0, hammer_s2_p: 0, hammer_s2_wr: '',
    hammer_s3_g: 0, hammer_s3_p: 0, hammer_s3_wr: '',

    tb_total_g: 0, tb_total_p: 0, tb_total_wr: '',
    tight_total_g: 0, tight_total_p: 0, tight_total_wr: '',
    cold_total_g: 0, cold_total_p: 0, cold_total_wr: '',
    hammer_total_g: 0, hammer_total_p: 0, hammer_total_wr: ''
  };
}

function processTypedSet(s, setType, setResult, slot) {
  if (!setType || !setResult) return;

  const typeMap = {
    TB: 'tb',
    TIGHT: 'tight',
    COLD: 'cold',
    HAMMER: 'hammer'
  };

  const prefix = typeMap[setType];
  if (!prefix) return;

  const gKey = `${prefix}_${slot}_g`;
  const pKey = `${prefix}_${slot}_p`;
  const totalGKey = `${prefix}_total_g`;
  const totalPKey = `${prefix}_total_p`;

  if (setResult === 'W') {
    s[gKey]++;
    s[totalGKey]++;
  } else if (setResult === 'L') {
    s[pKey]++;
    s[totalPKey]++;
  }
}

function computeTypedRates(s) {
  ['tb','tight','cold','hammer'].forEach(prefix => {
    ['s1','s2','s3'].forEach(slot => {
      const g = s[`${prefix}_${slot}_g`];
      const p = s[`${prefix}_${slot}_p`];
      s[`${prefix}_${slot}_wr`] = rate(g, g + p);
    });

    const tg = s[`${prefix}_total_g`];
    const tp = s[`${prefix}_total_p`];
    s[`${prefix}_total_wr`] = rate(tg, tg + tp);
  });
}
function buildStatsIndivProfile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = ss.getSheetByName('stats_indiv');

  if (!out) throw new Error('No existe la hoja stats_indiv');

  const data = out.getDataRange().getValues();
  if (data.length < 2) throw new Error('stats_indiv no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const requiredHeaders = [
    'tb_total_g', 'tb_total_p',
    'tight_total_g', 'tight_total_p',
    'cold_total_g', 'cold_total_p',
    'hammer_total_g', 'hammer_total_p',
    'pg_3s', 'pp_3s',
    'pg_2s', 'pp_2s'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan columnas en stats_indiv: ' + missing.join(', '));
  }

  const profileHeaders = [
    'hard_sets_g', 'hard_sets_p', 'hard_sets_wr',
    'easy_sets_g', 'easy_sets_p', 'easy_sets_wr',
    'long_m_g', 'long_m_p', 'long_m_wr',
    'short_m_g', 'short_m_p', 'short_m_wr'
  ];

  let currentLastCol = headers.length;
  profileHeaders.forEach(h => {
    if (!(h in idx)) {
      currentLastCol++;
      out.getRange(1, currentLastCol).setValue(h);
      idx[h] = currentLastCol - 1;
    }
  });

  const values = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];

    const tbG = Number(r[idx['tb_total_g']] || 0);
    const tbP = Number(r[idx['tb_total_p']] || 0);
    const tightG = Number(r[idx['tight_total_g']] || 0);
    const tightP = Number(r[idx['tight_total_p']] || 0);
    const coldG = Number(r[idx['cold_total_g']] || 0);
    const coldP = Number(r[idx['cold_total_p']] || 0);
    const hammerG = Number(r[idx['hammer_total_g']] || 0);
    const hammerP = Number(r[idx['hammer_total_p']] || 0);

    const pg3 = Number(r[idx['pg_3s']] || 0);
    const pp3 = Number(r[idx['pp_3s']] || 0);
    const pg2 = Number(r[idx['pg_2s']] || 0);
    const pp2 = Number(r[idx['pp_2s']] || 0);

    const hardG = tbG + tightG;
    const hardP = tbP + tightP;
    const easyG = coldG + hammerG;
    const easyP = coldP + hammerP;

    const longG = pg3;
    const longP = pp3;
    const shortG = pg2;
    const shortP = pp2;

    values.push([
      hardG,
      hardP,
      rate(hardG, hardG + hardP),
      easyG,
      easyP,
      rate(easyG, easyG + easyP),
      longG,
      longP,
      rate(longG, longG + longP),
      shortG,
      shortP,
      rate(shortG, shortG + shortP)
    ]);
  }

  if (values.length > 0) {
    out.getRange(2, idx['hard_sets_g'] + 1, values.length, profileHeaders.length).setValues(values);
  }

  ['hard_sets_g','hard_sets_p','easy_sets_g','easy_sets_p','long_m_g','long_m_p','short_m_g','short_m_p'].forEach(h => {
    const col = idx[h] + 1;
    out.getRange(2, col, out.getLastRow() - 1, 1).setNumberFormat('0');
  });

  ['hard_sets_wr','easy_sets_wr','long_m_wr','short_m_wr'].forEach(h => {
    const col = idx[h] + 1;
    out.getRange(2, col, out.getLastRow() - 1, 1).setNumberFormat('0.0%');
  });

  notifyUser(
    'stats_indiv Perfil reconstruido\n' +
    'Columnas agregadas/actualizadas: ' + profileHeaders.length
  );
}
function refreshAll() {

  try {
    SpreadsheetApp.flush();

assignPlayedOrder();
syncReplacementPlayersIntoPlayers();
syncLeaguePlayersFromMatches();

buildPlayerMatchLog();
    buildH2HPairs();
    buildH2HPlayerView();
    buildH2HMatchLog();

    buildCurrentPlayerStatus();

    buildRankingLast14();
    buildStatsRachas();

    buildStatsIndivBlock1();
    buildStatsIndivBlock2();
    buildStatsIndivBlock3();
    buildStatsIndivBlock4();
    buildStatsIndivProfile();

    buildDivisionStandings();
    buildDivisionFixtures();

    buildAppH2HSummary();
    buildAppH2HMatches();

    buildPlayerLeagueHistory();
buildPlayerRankSnapshots();
buildAppPlayers();
buildLeagueStatsSummary();
buildLeagueStatsRankings();

buildAppPlayerProfiles();

SpreadsheetApp.flush();

notifyUser('Modelo completo actualizado correctamente');
  } catch (err) {
    notifyUser('Error al actualizar el modelo: ' + err.message);
    throw err;
  }
}
function buildDivisionStandings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('matches_raw');
  const out = ss.getSheetByName('division_standings');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!out) throw new Error('No existe la hoja division_standings');

  const data = raw.getDataRange().getValues();
  if (data.length < 2) throw new Error('matches_raw no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const requiredHeaders = [
    'campeonato', 'division',
    'pareja1_j1', 'pareja1_j2', 'pareja2_j1', 'pareja2_j2',
    'winner', 'set1_p1', 'set1_p2', 'set2_p1', 'set2_p2', 'set3_p1', 'set3_p2',
    'result_type'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en matches_raw: ' + missing.join(', '));
  }

  const standings = {};

  rows.forEach(r => {
    const leagueId = safe(r[idx['campeonato']]);
    const division = safe(r[idx['division']]).toUpperCase();

    const p1a = safe(r[idx['pareja1_j1']]);
    const p1b = safe(r[idx['pareja1_j2']]);
    const p2a = safe(r[idx['pareja2_j1']]);
    const p2b = safe(r[idx['pareja2_j2']]);

    const winner = safe(r[idx['winner']]).toUpperCase();
    const resultType = safe(r[idx['result_type']]).toUpperCase();

    const isWO =
      resultType === 'WO_2_0' ||
      resultType === 'WO_INJURY_2_0' ||
      resultType === 'WO_INJURY_2_1';

    const s1a = toNumOrZero(r[idx['set1_p1']]);
    const s1b = toNumOrZero(r[idx['set1_p2']]);
    const s2a = toNumOrZero(r[idx['set2_p1']]);
    const s2b = toNumOrZero(r[idx['set2_p2']]);
    const s3aRaw = r[idx['set3_p1']];
    const s3bRaw = r[idx['set3_p2']];
    const s3a = toNumOrZero(s3aRaw);
    const s3b = toNumOrZero(s3bRaw);

    const hasPlayers =
      leagueId !== '' &&
      division !== '' &&
      p1a !== '' && p1b !== '' && p2a !== '' && p2b !== '';

    const hasWinner = winner === 'P1' || winner === 'P2';

    const hasRealScore =
      r[idx['set1_p1']] !== '' && r[idx['set1_p1']] !== null &&
      r[idx['set1_p2']] !== '' && r[idx['set1_p2']] !== null &&
      r[idx['set2_p1']] !== '' && r[idx['set2_p1']] !== null &&
      r[idx['set2_p2']] !== '' && r[idx['set2_p2']] !== null;

    const played = hasPlayers && hasWinner && (hasRealScore || isWO);

    if (!played) return;

    const leagueDivisionId = `${leagueId}_${division}`;

    if (!standings[leagueDivisionId]) {
      standings[leagueDivisionId] = {};
    }

    const group = standings[leagueDivisionId];

    [p1a, p1b, p2a, p2b].forEach(player => {
      if (!group[player]) {
        group[player] = {
          league_division_id: leagueDivisionId,
          league_id: leagueId,
          division: division,
          player: player,
          pj: 0,
          pg: 0,
          pp: 0,
          pts: 0,
          sets_g: 0,
          sets_p: 0,
          games_g: 0,
          games_p: 0
        };
      }
    });

    let sets1 = 0, sets2 = 0;
    let games1 = 0, games2 = 0;

    if (isWO) {
      // Resultado administrativo: suma puntos y sets para tabla, pero no games.
      if (winner === 'P1') {
        sets1 = 2;
        sets2 = resultType === 'WO_INJURY_2_1' ? 1 : 0;
      } else {
        sets2 = 2;
        sets1 = resultType === 'WO_INJURY_2_1' ? 1 : 0;
      }
    } else {
      games1 = s1a + s2a + s3a;
      games2 = s1b + s2b + s3b;

      if (s1a > s1b) sets1++; else sets2++;
      if (s2a > s2b) sets1++; else sets2++;

      const hasThird = (s3aRaw !== '' && s3aRaw !== null && s3bRaw !== '' && s3bRaw !== null);
      if (hasThird) {
        if (s3a > s3b) sets1++;
        else sets2++;
      }
    }

    let pts1 = 0, pts2 = 0;
    let won1 = 0, won2 = 0;
    let lost1 = 0, lost2 = 0;

    if (winner === 'P1') {
      won1 = 1;
      lost2 = 1;

      if (isWO) {
        if (resultType === 'WO_INJURY_2_1') {
          pts1 = 2;
          pts2 = 1;
        } else {
          pts1 = 3;
          pts2 = 0;
        }
      } else if (sets1 === 2 && sets2 === 0) {
        pts1 = 3;
        pts2 = 0;
      } else {
        pts1 = 2;
        pts2 = 1;
      }
    } else {
      won2 = 1;
      lost1 = 1;

      if (isWO) {
        if (resultType === 'WO_INJURY_2_1') {
          pts2 = 2;
          pts1 = 1;
        } else {
          pts2 = 3;
          pts1 = 0;
        }
      } else if (sets2 === 2 && sets1 === 0) {
        pts2 = 3;
        pts1 = 0;
      } else {
        pts2 = 2;
        pts1 = 1;
      }
    }

    [p1a, p1b].forEach(player => {
      group[player].pj += 1;
      group[player].pg += won1;
      group[player].pp += lost1;
      group[player].pts += pts1;
      group[player].sets_g += sets1;
      group[player].sets_p += sets2;
      group[player].games_g += games1;
      group[player].games_p += games2;
    });

    [p2a, p2b].forEach(player => {
      group[player].pj += 1;
      group[player].pg += won2;
      group[player].pp += lost2;
      group[player].pts += pts2;
      group[player].sets_g += sets2;
      group[player].sets_p += sets1;
      group[player].games_g += games2;
      group[player].games_p += games1;
    });
  });

  const outRows = [];

  Object.keys(standings).sort().forEach(leagueDivisionId => {
    const players = Object.values(standings[leagueDivisionId]);

    players.sort((a, b) => {
      if (b.pts !== a.pts) return b.pts - a.pts;

      const diffA = a.games_g - a.games_p;
      const diffB = b.games_g - b.games_p;
      if (diffB !== diffA) return diffB - diffA;

      if (b.pg !== a.pg) return b.pg - a.pg;

      const setDiffA = a.sets_g - a.sets_p;
      const setDiffB = b.sets_g - b.sets_p;
      if (setDiffB !== setDiffA) return setDiffB - setDiffA;

      return a.player.localeCompare(b.player);
    });

    players.forEach((p, i) => {
      outRows.push([
        p.league_division_id,
        p.league_id,
        p.division,
        i + 1,
        p.player,
        p.pj,
        p.pg,
        p.pp,
        p.pts,
        p.sets_g,
        p.sets_p,
        p.games_g,
        p.games_p,
        p.games_g - p.games_p
      ]);
    });
  });

  out.clearContents();

  out.getRange(1, 1, 1, 14).setValues([[
    'league_division_id','league_id','division','pos','player','pj','pg','pp','pts','sets_g','sets_p','games_g','games_p','diff_games'
  ]]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, 14).setValues(outRows);
  }

  notifyUser(
    'division_standings reconstruida\n' +
    'Filas generadas: ' + outRows.length
  );
}
function buildDivisionFixtures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('matches_raw');
  const out = ss.getSheetByName('division_fixtures');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!out) throw new Error('No existe la hoja division_fixtures');

  const data = raw.getDataRange().getValues();
  if (data.length < 2) throw new Error('matches_raw no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const requiredHeaders = [
    'match_id', 'campeonato', 'fecha', 'division',
    'pareja1_j1', 'pareja1_j2', 'pareja2_j1', 'pareja2_j2',
    'winner', 'set1_p1', 'set1_p2', 'set2_p1', 'set2_p2', 'set3_p1', 'set3_p2',
    'result_type'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan encabezados en matches_raw: ' + missing.join(', '));
  }

  const outRows = [];

  rows.forEach(r => {
    const matchId = safe(r[idx['match_id']]);
    const leagueId = safe(r[idx['campeonato']]);
    const division = safe(r[idx['division']]).toUpperCase();
    const fecha = r[idx['fecha']];

    const p1j1 = safe(r[idx['pareja1_j1']]);
    const p1j2 = safe(r[idx['pareja1_j2']]);
    const p2j1 = safe(r[idx['pareja2_j1']]);
    const p2j2 = safe(r[idx['pareja2_j2']]);

    const winner = safe(r[idx['winner']]).toUpperCase();
    const resultType = safe(r[idx['result_type']]).toUpperCase();

    const isWO =
      resultType === 'WO_2_0' ||
      resultType === 'WO_INJURY_2_0' ||
      resultType === 'WO_INJURY_2_1';

    const s1p1Raw = r[idx['set1_p1']];
    const s1p2Raw = r[idx['set1_p2']];
    const s2p1Raw = r[idx['set2_p1']];
    const s2p2Raw = r[idx['set2_p2']];
    const s3p1Raw = r[idx['set3_p1']];
    const s3p2Raw = r[idx['set3_p2']];

    const hasPlayers =
      leagueId !== '' &&
      division !== '' &&
      p1j1 !== '' && p1j2 !== '' && p2j1 !== '' && p2j2 !== '';

    const hasWinner = winner === 'P1' || winner === 'P2';

    const hasRealScore =
      s1p1Raw !== '' && s1p1Raw !== null &&
      s1p2Raw !== '' && s1p2Raw !== null &&
      s2p1Raw !== '' && s2p1Raw !== null &&
      s2p2Raw !== '' && s2p2Raw !== null;

    const played = hasPlayers && hasWinner && (hasRealScore || isWO);

    const leagueDivisionId = `${leagueId}_${division}`;
    const pair1Display = `${p1j1} / ${p1j2}`;
    const pair2Display = `${p2j1} / ${p2j2}`;

    let matchStatus = 'SCHEDULED';
    let matchFormat = '';
    let scoreDisplay = '';
    let setsP1 = '';
    let setsP2 = '';
    let gamesP1 = '';
    let gamesP2 = '';
    let pointsP1 = '';
    let pointsP2 = '';

    if (played) {
      if (isWO) {
        matchStatus = 'WO';
        matchFormat = resultType === 'WO_INJURY_2_1' ? '3S' : '2S';

        if (winner === 'P1') {
          scoreDisplay = resultType === 'WO_INJURY_2_1' ? 'WO 2-1' : 'WO 2-0';
          setsP1 = 2;
          setsP2 = resultType === 'WO_INJURY_2_1' ? 1 : 0;
          gamesP1 = 0;
          gamesP2 = 0;
          pointsP1 = resultType === 'WO_INJURY_2_1' ? 2 : 3;
          pointsP2 = resultType === 'WO_INJURY_2_1' ? 1 : 0;
        } else {
          scoreDisplay = resultType === 'WO_INJURY_2_1' ? 'WO 1-2' : 'WO 0-2';
          setsP1 = resultType === 'WO_INJURY_2_1' ? 1 : 0;
          setsP2 = 2;
          gamesP1 = 0;
          gamesP2 = 0;
          pointsP1 = resultType === 'WO_INJURY_2_1' ? 1 : 0;
          pointsP2 = resultType === 'WO_INJURY_2_1' ? 2 : 3;
        }
      } else {
        matchStatus = 'PLAYED';

        const s1p1 = toNumOrZero(s1p1Raw);
        const s1p2 = toNumOrZero(s1p2Raw);
        const s2p1 = toNumOrZero(s2p1Raw);
        const s2p2 = toNumOrZero(s2p2Raw);
        const s3p1 = toNumOrZero(s3p1Raw);
        const s3p2 = toNumOrZero(s3p2Raw);

        const setPairs = [
          [s1p1, s1p2],
          [s2p1, s2p2]
        ];

        const hasThird =
          s3p1Raw !== '' && s3p1Raw !== null &&
          s3p2Raw !== '' && s3p2Raw !== null;

        if (hasThird) {
          setPairs.push([s3p1, s3p2]);
          matchFormat = '3S';
        } else {
          matchFormat = '2S';
        }

        scoreDisplay = setPairs.map(([a, b]) => `${a}-${b}`).join(', ');

        let sp1 = 0;
        let sp2 = 0;

        setPairs.forEach(([a, b]) => {
          if (a > b) sp1++;
          else if (b > a) sp2++;
        });

        setsP1 = sp1;
        setsP2 = sp2;
        gamesP1 = setPairs.reduce((acc, [a]) => acc + a, 0);
        gamesP2 = setPairs.reduce((acc, [, b]) => acc + b, 0);

        if (winner === 'P1') {
          if (sp1 === 2 && sp2 === 0) {
            pointsP1 = 3;
            pointsP2 = 0;
          } else {
            pointsP1 = 2;
            pointsP2 = 1;
          }
        } else if (winner === 'P2') {
          if (sp2 === 2 && sp1 === 0) {
            pointsP1 = 0;
            pointsP2 = 3;
          } else {
            pointsP1 = 1;
            pointsP2 = 2;
          }
        }
      }
    }

    outRows.push([
      leagueDivisionId,
      leagueId,
      division,
      fecha,
      matchId,
      p1j1,
      p1j2,
      p2j1,
      p2j2,
      pair1Display,
      pair2Display,
      winner,
      matchStatus,
      matchFormat,
      scoreDisplay,
      setsP1,
      setsP2,
      gamesP1,
      gamesP2,
      pointsP1,
      pointsP2
    ]);
  });

  outRows.sort((a, b) => {
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;

    const fa = String(a[3]);
    const fb = String(b[3]);
    if (fa < fb) return -1;
    if (fa > fb) return 1;

    const ma = String(a[4]);
    const mb = String(b[4]);
    return ma.localeCompare(mb);
  });

  out.clearContents();

  out.getRange(1, 1, 1, 21).setValues([[
    'league_division_id','league_id','division','fecha','match_id',
    'pair1_j1','pair1_j2','pair2_j1','pair2_j2',
    'pair1_display','pair2_display',
    'winner','match_status','match_format','score_display',
    'sets_p1','sets_p2','games_p1','games_p2','points_p1','points_p2'
  ]]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, 21).setValues(outRows);
  }

  notifyUser(
    'division_fixtures reconstruida\n' +
    'Filas generadas: ' + outRows.length
  );
}
function buildAppH2HSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseH2H = ss.getSheetByName('h2h_player_view');
  const players = ss.getSheetByName('players');
  const ranking = ss.getSheetByName('ranking_last14');
  const currentStatus = ss.getSheetByName('current_player_status');
  const rachas = ss.getSheetByName('stats_rachas');
  const out = ss.getSheetByName('app_h2h_summary');

  if (!baseH2H) throw new Error('No existe la hoja h2h_player_view');
  if (!players) throw new Error('No existe la hoja players');
  if (!ranking) throw new Error('No existe la hoja ranking_last14');
  if (!currentStatus) throw new Error('No existe la hoja current_player_status');
  if (!rachas) throw new Error('No existe la hoja stats_rachas');
  if (!out) throw new Error('No existe la hoja app_h2h_summary');

  const h2hData = baseH2H.getDataRange().getValues();
  const playersData = players.getDataRange().getValues();
  const rankingData = ranking.getDataRange().getValues();
  const statusData = currentStatus.getDataRange().getValues();
  const rachasData = rachas.getDataRange().getValues();

  if (h2hData.length < 2) throw new Error('h2h_player_view no tiene datos');

  const h2hHeaders = h2hData[0].map(h => String(h).trim());
  const h2hRows = h2hData.slice(1);

  const pHeaders = playersData[0].map(h => String(h).trim());
  const pRows = playersData.slice(1);

  const rankHeaders = rankingData[0].map(h => String(h).trim());
  const rankRows = rankingData.slice(1);

  const statusHeaders = statusData[0].map(h => String(h).trim());
  const statusRows = statusData.slice(1);

  const rHeaders = rachasData[0].map(h => String(h).trim());
  const rRows = rachasData.slice(1);

  const h2hIdx = {};
  h2hHeaders.forEach((h, i) => h2hIdx[h] = i);

  const pIdx = {};
  pHeaders.forEach((h, i) => pIdx[h] = i);

  const rankIdx = {};
  rankHeaders.forEach((h, i) => rankIdx[h] = i);

  const statusIdx = {};
  statusHeaders.forEach((h, i) => statusIdx[h] = i);

  const rIdx = {};
  rHeaders.forEach((h, i) => rIdx[h] = i);

  const requiredH2H = [
    'player_a','player_b',
    'matches_a','matches_b','matches_wr_a','matches_wr_b',
    'sets_a','sets_b','sets_wr_a','sets_wr_b',
    'games_a','games_b','games_wr_a','games_wr_b'
  ];

  const missing = requiredH2H.filter(h => !(h in h2hIdx));
  if (missing.length > 0) {
    throw new Error('Faltan columnas en h2h_player_view: ' + missing.join(', '));
  }

  const playersMap = {};
  pRows.forEach(r => {
    const playerName = safe(r[pIdx['player_name']]);
    if (!playerName) return;
    playersMap[playerName] = {
      photo_url: ('photo_url' in pIdx) ? safe(r[pIdx['photo_url']]) : '',
      display_name: ('display_name' in pIdx) ? safe(r[pIdx['display_name']]) : playerName
    };
  });

  const rankingMap = {};
  rankRows.forEach(r => {
    const player = safe(r[rankIdx['player']]);
    if (!player) return;
    rankingMap[player] = {
      rank_position: ('rank_position' in rankIdx) ? r[rankIdx['rank_position']] : ''
    };
  });

  const statusMap = {};
  statusRows.forEach(r => {
    const player = safe(r[statusIdx['player']]);
    if (!player) return;
    statusMap[player] = {
      division_actual: ('division_actual' in statusIdx) ? safe(r[statusIdx['division_actual']]) : '-'
    };
  });

  const rachasMap = {};
  rRows.forEach(r => {
    const player = safe(r[rIdx['player']]);
    if (!player) return;
    rachasMap[player] = {
      racha_actual: ('racha_actual' in rIdx) ? safe(r[rIdx['racha_actual']]) : ''
    };
  });

  const outRows = [];

  h2hRows.forEach(r => {
    const playerA = safe(r[h2hIdx['player_a']]);
    const playerB = safe(r[h2hIdx['player_b']]);
    if (!playerA || !playerB) return;

    const matchesA = Number(r[h2hIdx['matches_a']] || 0);
    const matchesB = Number(r[h2hIdx['matches_b']] || 0);
    const setsA = Number(r[h2hIdx['sets_a']] || 0);
    const setsB = Number(r[h2hIdx['sets_b']] || 0);
    const gamesA = Number(r[h2hIdx['games_a']] || 0);
    const gamesB = Number(r[h2hIdx['games_b']] || 0);

    const totalMatches = matchesA + matchesB;
    const totalSets = setsA + setsB;
    const totalGames = gamesA + gamesB;

    outRows.push([
      playerA,
      playerB,
      'Todas',
      'Todas',

      playersMap[playerA]?.photo_url || '',
      playersMap[playerA]?.display_name || playerA,
      statusMap[playerA]?.division_actual || '-',
      rankingMap[playerA]?.rank_position || '',
      rachasMap[playerA]?.racha_actual || '',

      playersMap[playerB]?.photo_url || '',
      playersMap[playerB]?.display_name || playerB,
      statusMap[playerB]?.division_actual || '-',
      rankingMap[playerB]?.rank_position || '',
      rachasMap[playerB]?.racha_actual || '',

      totalMatches,
      matchesA,
      matchesB,
      totalMatches > 0 ? matchesA / totalMatches : '',
      totalMatches > 0 ? matchesB / totalMatches : '',

      totalSets,
      setsA,
      setsB,
      totalSets > 0 ? setsA / totalSets : '',
      totalSets > 0 ? setsB / totalSets : '',

      totalGames,
      gamesA,
      gamesB,
      totalGames > 0 ? gamesA / totalGames : '',
      totalGames > 0 ? gamesB / totalGames : ''
    ]);
  });

  out.clearContents();
  out.getRange(1, 1, 1, 29).setValues([[
    'player_a','player_b','temporada','division',
    'player_a_photo_url','player_a_display_name','player_a_division_actual','player_a_rank_position','player_a_racha_actual',
    'player_b_photo_url','player_b_display_name','player_b_division_actual','player_b_rank_position','player_b_racha_actual',
    'total_matches','matches_a','matches_b','matches_wr_a','matches_wr_b',
    'total_sets','sets_a','sets_b','sets_wr_a','sets_wr_b',
    'total_games','games_a','games_b','games_wr_a','games_wr_b'
  ]]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, 29).setValues(outRows);
    [18,19,23,24,28,29].forEach(col => {
      out.getRange(2, col, outRows.length, 1).setNumberFormat('0.0%');
    });
  }

  notifyUser(
    'app_h2h_summary reconstruida\n' +
    'Filas generadas: ' + outRows.length
  );
}
function buildAppH2HMatches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName('h2h_match_log');
  const out = ss.getSheetByName('app_h2h_matches');

  if (!base) throw new Error('No existe la hoja h2h_match_log');
  if (!out) throw new Error('No existe la hoja app_h2h_matches');

  const data = base.getDataRange().getValues();
  if (data.length < 2) throw new Error('h2h_match_log no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const requiredHeaders = [
    'player_a', 'player_b',
    'campeonato', 'fecha', 'division',
    'match_id', 'chronological_index',
    'pair_a_display', 'pair_b_display',
    'score_display', 'winner_h2h',
    'player_a_side'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan columnas en h2h_match_log: ' + missing.join(', '));
  }

  const outRows = [];

  rows.forEach(r => {
    const playerA = safe(r[idx['player_a']]);
    const playerB = safe(r[idx['player_b']]);
    const temporada = safe(r[idx['campeonato']]);
    const fechaRaw = safe(r[idx['fecha']]);
    const division = safe(r[idx['division']]).toUpperCase();
    const matchId = safe(r[idx['match_id']]);
    const chronologicalIndex = Number(r[idx['chronological_index']] || 0);
    const pairADisplay = safe(r[idx['pair_a_display']]);
    const pairBDisplay = safe(r[idx['pair_b_display']]);
    const scoreDisplayNeutral = safe(r[idx['score_display']]);
    const winnerH2H = safe(r[idx['winner_h2h']]);
    const playerASide = safe(r[idx['player_a_side']]).toUpperCase();

    if (!playerA || !playerB || !temporada || !matchId) return;

    const fechaLiga = extractFechaLiga(fechaRaw);
    const scoreDisplayPerspective = orientScoreToPlayerA(scoreDisplayNeutral, playerASide);

    outRows.push([
      playerA,
      playerB,
      temporada,
      fechaLiga,
      division,
      matchId,
      chronologicalIndex,
      pairADisplay,
      pairBDisplay,
      scoreDisplayPerspective,
      winnerH2H
    ]);
  });

  outRows.sort((a, b) => {
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;
    if (a[1] < b[1]) return -1;
    if (a[1] > b[1]) return 1;
    return b[6] - a[6];
  });

  out.clearContents();
  out.getRange(1, 1, 1, 11).setValues([[
    'player_a','player_b','temporada','fecha_liga','division','match_id','chronological_index',
    'pair_a_display','pair_b_display','score_display_perspective','winner_h2h'
  ]]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, 11).setValues(outRows);
  }

  notifyUser(
    'app_h2h_matches reconstruida\n' +
    'Filas generadas: ' + outRows.length
  );
}

function orientScoreToPlayerA(scoreDisplay, playerASide) {
  if (!scoreDisplay) return '';
  if (playerASide === 'P1') return scoreDisplay;

  const sets = scoreDisplay.split(',').map(s => s.trim()).filter(Boolean);

  const flipped = sets.map(setText => {
    const parts = setText.split('-').map(x => x.trim());
    if (parts.length !== 2) return setText;
    return `${parts[1]}-${parts[0]}`;
  });

  return flipped.join(', ');
}

function extractFechaLiga(fechaRaw) {
  if (!fechaRaw) return '';

  const txt = String(fechaRaw).trim();

  // Si ya es un número puro
  if (/^\d+$/.test(txt)) return Number(txt);

  // Busca algo como "Fecha 1", "F1", "fecha 7", etc.
  const match = txt.match(/(\d+)/);
  if (match) return Number(match[1]);

  return txt;
}
function buildPlayerLeagueHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const standings = ss.getSheetByName('division_standings');
  const out = ss.getSheetByName('player_league_history');

  if (!standings) throw new Error('No existe la hoja division_standings');
  if (!out) throw new Error('No existe la hoja player_league_history');

  const data = standings.getDataRange().getValues();
  if (data.length < 2) throw new Error('division_standings no tiene datos');

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const requiredHeaders = [
    'league_id', 'division', 'pos', 'player',
    'pj', 'pg', 'pp', 'pts', 'diff_games'
  ];

  const missing = requiredHeaders.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan columnas en division_standings: ' + missing.join(', '));
  }

  const outRows = rows
    .filter(r => safe(r[idx['player']]) !== '')
    .map(r => {
      const player = safe(r[idx['player']]);
      const leagueId = safe(r[idx['league_id']]);
      const division = safe(r[idx['division']]).toUpperCase();
      const finalPos = Number(r[idx['pos']] || 0);
      const pj = Number(r[idx['pj']] || 0);
      const pg = Number(r[idx['pg']] || 0);
      const pp = Number(r[idx['pp']] || 0);
      const pts = Number(r[idx['pts']] || 0);
      const diffGames = Number(r[idx['diff_games']] || 0);

      return [
        player,
        leagueId,
        division,
        finalPos,
        pj,
        pg,
        pp,
        pts,
        diffGames,
        buildLeagueHistoryLabel(finalPos, division, leagueId),
        finalPos === 1 ? 1 : 0,
        finalPos <= 3 ? 1 : 0
      ];
    });

  outRows.sort((a, b) => {
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;
    if (a[1] < b[1]) return -1;
    if (a[1] > b[1]) return 1;
    if (a[2] < b[2]) return -1;
    if (a[2] > b[2]) return 1;
    return a[3] - b[3];
  });

  out.clearContents();
  out.getRange(1, 1, 1, 12).setValues([[
    'player','league_id','division','final_pos','pj','pg','pp','pts','diff_games','label_result','is_champion','is_podium'
  ]]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, 12).setValues(outRows);
  }

  notifyUser(
    'player_league_history reconstruida\n' +
    'Filas generadas: ' + outRows.length
  );
}

function buildLeagueHistoryLabel(pos, division, leagueId) {
  if (pos === 1) {
    return `Campeón División ${division} — ${leagueId}`;
  }
  if (pos === 2) {
    return `2° lugar División ${division} — ${leagueId}`;
  }
  if (pos === 3) {
    return `3° lugar División ${division} — ${leagueId}`;
  }
  return `${pos}° lugar División ${division} — ${leagueId}`;
}
function buildPlayerRankSnapshots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('matches_raw');
  const snaps = ss.getSheetByName('league_snapshots');
  const out = ss.getSheetByName('player_rank_snapshots');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!snaps) throw new Error('No existe la hoja league_snapshots');
  if (!out) throw new Error('No existe la hoja player_rank_snapshots');

  const rawData = raw.getDataRange().getValues();
  const snapData = snaps.getDataRange().getValues();

  if (rawData.length < 2) throw new Error('matches_raw no tiene datos');
  if (snapData.length < 2) throw new Error('league_snapshots no tiene datos');

  const rawHeaders = rawData[0].map(h => String(h).trim());
  const rawRows = rawData.slice(1);

  const snapHeaders = snapData[0].map(h => String(h).trim());
  const snapRows = snapData.slice(1);

  const rawIdx = {};
  rawHeaders.forEach((h, i) => rawIdx[h] = i);

  const snapIdx = {};
  snapHeaders.forEach((h, i) => snapIdx[h] = i);

  const requiredRaw = [
    'campeonato', 'division',
    'pareja1_j1', 'pareja1_j2', 'pareja2_j1', 'pareja2_j2',
    'winner', 'set1_p1', 'set1_p2', 'set2_p1', 'set2_p2', 'set3_p1', 'set3_p2'
  ];

  const missingRaw = requiredRaw.filter(h => !(h in rawIdx));
  if (missingRaw.length > 0) {
    throw new Error('Faltan columnas en matches_raw: ' + missingRaw.join(', '));
  }

  const requiredSnap = [
    'snapshot_order', 'snapshot_label', 'cutoff_league_id', 'is_active'
  ];

  const missingSnap = requiredSnap.filter(h => !(h in snapIdx));
  if (missingSnap.length > 0) {
    throw new Error('Faltan columnas en league_snapshots: ' + missingSnap.join(', '));
  }

  const snapshots = snapRows
    .filter(r => {
      const active = String(r[snapIdx['is_active']]).trim().toLowerCase();
      return active === '1' || active === 'true';
    })
    .map(r => ({
      snapshot_order: Number(r[snapIdx['snapshot_order']] || 0),
      snapshot_label: safe(r[snapIdx['snapshot_label']]),
      cutoff_league_id: safe(r[snapIdx['cutoff_league_id']])
    }))
    .sort((a, b) => a.snapshot_order - b.snapshot_order);

  if (snapshots.length === 0) {
  out.clearContents();
  out.getRange(1, 1, 1, 9).setValues([[
    'snapshot_label',
    'snapshot_order',
    'cutoff_league_id',
    'player',
    'matches_available',
    'matches_counted',
    'score_snapshot',
    'avg_snapshot',
    'rank_position'
  ]]);

  notifyUser(
    'player_rank_snapshots sin snapshots activos. Se dejó solo encabezado.'
  );

  return;
}

  const playerMatches = {};
  const generatedSnapshots = new Set();
  const outRows = [];

  for (let i = 0; i < rawRows.length; i++) {
    const r = rawRows[i];

    const leagueId = safe(r[rawIdx['campeonato']]);
    const division = safe(r[rawIdx['division']]).toUpperCase();

    const p1a = safe(r[rawIdx['pareja1_j1']]);
    const p1b = safe(r[rawIdx['pareja1_j2']]);
    const p2a = safe(r[rawIdx['pareja2_j1']]);
    const p2b = safe(r[rawIdx['pareja2_j2']]);
    const winner = safe(r[rawIdx['winner']]).toUpperCase();

    const s1aRaw = r[rawIdx['set1_p1']];
    const s1bRaw = r[rawIdx['set1_p2']];
    const s2aRaw = r[rawIdx['set2_p1']];
    const s2bRaw = r[rawIdx['set2_p2']];
    const s3aRaw = r[rawIdx['set3_p1']];
    const s3bRaw = r[rawIdx['set3_p2']];

    const hasPlayers =
  leagueId !== '' &&
  division !== '' &&
  p1a !== '' && p1b !== '' && p2a !== '' && p2b !== '';

const hasWinner = winner === 'P1' || winner === 'P2';

const hasRealScore =
  s1aRaw !== '' && s1aRaw !== null &&
  s1bRaw !== '' && s1bRaw !== null &&
  s2aRaw !== '' && s2aRaw !== null &&
  s2bRaw !== '' && s2bRaw !== null;

const played = hasPlayers && hasWinner && hasRealScore;

    if (!played) continue;

    const s1a = toNumOrZero(s1aRaw);
    const s1b = toNumOrZero(s1bRaw);
    const s2a = toNumOrZero(s2aRaw);
    const s2b = toNumOrZero(s2bRaw);
    const s3a = toNumOrZero(s3aRaw);
    const s3b = toNumOrZero(s3bRaw);

    let sets1 = 0, sets2 = 0;
    const games1 = s1a + s2a + s3a;
    const games2 = s1b + s2b + s3b;

    if (s1a > s1b) sets1++; else sets2++;
    if (s2a > s2b) sets1++; else sets2++;

    if ((s3a !== 0 || s3b !== 0)) {
      if (s3a > s3b) sets1++;
      else sets2++;
    }

    let res1 = '', res2 = '';

    if (winner === 'P1') {
      if (sets1 === 2 && sets2 === 0) {
        res1 = 'G2'; res2 = 'P2';
      } else {
        res1 = 'G3'; res2 = 'P3';
      }
    } else {
      if (sets2 === 2 && sets1 === 0) {
        res1 = 'P2'; res2 = 'G2';
      } else {
        res1 = 'P3'; res2 = 'G3';
      }
    }

    const pts1 = rankPointsOfficial(division, res1, games1 - games2);
    const pts2 = rankPointsOfficial(division, res2, games2 - games1);

    addSnapshotMatch(playerMatches, p1a, pts1);
    addSnapshotMatch(playerMatches, p1b, pts1);
    addSnapshotMatch(playerMatches, p2a, pts2);
    addSnapshotMatch(playerMatches, p2b, pts2);

    const nextLeagueId = (i < rawRows.length - 1) ? safe(rawRows[i + 1][rawIdx['campeonato']]) : '';

    // Si la próxima fila ya es de otra liga, entonces esta fue la última fila de leagueId
    const isEndOfLeague = (leagueId !== nextLeagueId);

    if (isEndOfLeague) {
      snapshots.forEach(snap => {
        if (snap.cutoff_league_id === leagueId && !generatedSnapshots.has(snap.snapshot_label)) {
          const snapshotRows = buildSnapshotRows(playerMatches, snap);
          snapshotRows.forEach(row => outRows.push(row));
          generatedSnapshots.add(snap.snapshot_label);
        }
      });
    }
  }

  out.clearContents();
  out.getRange(1, 1, 1, 9).setValues([[
    'snapshot_label','snapshot_order','cutoff_league_id','player',
    'matches_available','matches_counted','score_snapshot','avg_snapshot','rank_position'
  ]]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, 9).setValues(outRows);
    out.getRange(2, 7, outRows.length, 1).setNumberFormat('0.0');
    out.getRange(2, 8, outRows.length, 1).setNumberFormat('0.00');
  }

  notifyUser(
    'player_rank_snapshots reconstruida\n' +
    'Snapshots generados: ' + generatedSnapshots.size + '\n' +
    'Filas generadas: ' + outRows.length
  );
}

function addSnapshotMatch(playerMatches, player, points) {
  if (!player) return;
  if (!playerMatches[player]) playerMatches[player] = [];
  playerMatches[player].push(points);
}

function buildSnapshotRows(playerMatches, snap) {
  const rows = [];

  Object.keys(playerMatches).forEach(player => {
    const matches = playerMatches[player];
    const matchesAvailable = matches.length;
    const matchesCounted = Math.min(14, matchesAvailable);
    const slice = matches.slice(-matchesCounted);

    let score = 0;
    slice.forEach(v => score += v);

   const scoreScaled = score * 10;
const avgScaled = matchesCounted > 0 ? scoreScaled / matchesCounted : 0;

rows.push({
  snapshot_label: snap.snapshot_label,
  snapshot_order: snap.snapshot_order,
  cutoff_league_id: snap.cutoff_league_id,
  player: player,
  matches_available: matchesAvailable,
  matches_counted: matchesCounted,
  score_snapshot: round1(scoreScaled),
  avg_snapshot: round2(avgScaled),
  rank_position: 0
});
  });

  rows.sort((a, b) => {
    if (b.score_snapshot !== a.score_snapshot) return b.score_snapshot - a.score_snapshot;
    if (b.avg_snapshot !== a.avg_snapshot) return b.avg_snapshot - a.avg_snapshot;
    return a.player.localeCompare(b.player);
  });

  rows.forEach((r, i) => {
    r.rank_position = i + 1;
  });

  return rows.map(r => [
    r.snapshot_label,
    r.snapshot_order,
    r.cutoff_league_id,
    r.player,
    r.matches_available,
    r.matches_counted,
    r.score_snapshot,
    r.avg_snapshot,
    r.rank_position
  ]);
}
function buildAppPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const players = ss.getSheetByName('players');
const ranking = ss.getSheetByName('ranking_last14');
const currentStatus = ss.getSheetByName('current_player_status');
const rachas = ss.getSheetByName('stats_rachas');
const indiv = ss.getSheetByName('stats_indiv');
const out = ss.getSheetByName('app_players');
if (!currentStatus) throw new Error('No existe la hoja current_player_status');

  if (!players) throw new Error('No existe la hoja players');
  if (!ranking) throw new Error('No existe la hoja ranking_last14');
  if (!rachas) throw new Error('No existe la hoja stats_rachas');
  if (!indiv) throw new Error('No existe la hoja stats_indiv');
  if (!out) throw new Error('No existe la hoja app_players');

  const pData = players.getDataRange().getValues();
const rankData = ranking.getDataRange().getValues();
const statusData = currentStatus.getDataRange().getValues();
const rData = rachas.getDataRange().getValues();
const iData = indiv.getDataRange().getValues();

  if (pData.length < 2) throw new Error('players no tiene datos');

  const pHeaders = pData[0].map(h => String(h).trim());
const rankHeaders = rankData[0].map(h => String(h).trim());
const statusHeaders = statusData[0].map(h => String(h).trim());
const rHeaders = rData[0].map(h => String(h).trim());
const iHeaders = iData[0].map(h => String(h).trim());

const pRows = pData.slice(1);
const rankRows = rankData.slice(1);
const statusRows = statusData.slice(1);
const rRows = rData.slice(1);
const iRows = iData.slice(1);

  const pIdx = {};
pHeaders.forEach((h, i) => pIdx[h] = i);

const rankIdx = {};
rankHeaders.forEach((h, i) => rankIdx[h] = i);

const statusIdx = {};
statusHeaders.forEach((h, i) => statusIdx[h] = i);

const rIdx = {};
rHeaders.forEach((h, i) => rIdx[h] = i);

const iIdx = {};
iHeaders.forEach((h, i) => iIdx[h] = i);

  const rankMap = {};
  rankRows.forEach(r => {
    const player = safe(r[rankIdx['player']]);
    if (!player) return;
    rankMap[player] = {
      division_actual: safe(r[rankIdx['division_actual']]),
      rank_position: ('rank_position' in rankIdx) ? r[rankIdx['rank_position']] : '',
      score_last14: ('score_last14' in rankIdx) ? r[rankIdx['score_last14']] : '',
      avg_last14: ('avg_last14' in rankIdx) ? r[rankIdx['avg_last14']] : ''
    };
  });
  const statusMap = {};
statusRows.forEach(r => {
  const player = safe(r[statusIdx['player']]);
  if (!player) return;
  statusMap[player] = {
    division_actual: ('division_actual' in statusIdx) ? safe(r[statusIdx['division_actual']]) : '-'
  };
});

  const rachasMap = {};
  rRows.forEach(r => {
    const player = safe(r[rIdx['player']]);
    if (!player) return;
    rachasMap[player] = {
      racha_actual: ('racha_actual' in rIdx) ? safe(r[rIdx['racha_actual']]) : '',
      mejor_racha_g: ('mejor_racha_g' in rIdx) ? r[rIdx['mejor_racha_g']] : '',
      mejor_racha_p: ('mejor_racha_p' in rIdx) ? r[rIdx['mejor_racha_p']] : '',
      seq_5: ('seq_5' in rIdx) ? safe(r[rIdx['seq_5']]) : '',
      seq_10: ('seq_10' in rIdx) ? safe(r[rIdx['seq_10']]) : ''
    };
  });

  const indivMap = {};
  iRows.forEach(r => {
    const player = safe(r[iIdx['player']]);
    if (!player) return;
    indivMap[player] = {
      match_wr: getVal(r, iIdx, 'match_wr'),
      sets_wr: getVal(r, iIdx, 'sets_wr'),
      games_wr: getVal(r, iIdx, 'games_wr'),
      wr_2s: getVal(r, iIdx, 'wr_2s'),
      wr_3s: getVal(r, iIdx, 'wr_3s'),
      hard_sets_wr: getVal(r, iIdx, 'hard_sets_wr'),
      easy_sets_wr: getVal(r, iIdx, 'easy_sets_wr'),
      long_m_wr: getVal(r, iIdx, 'long_m_wr'),
      short_m_wr: getVal(r, iIdx, 'short_m_wr'),
      cbs_total: getVal(r, iIdx, 'cbs_total'),
      farra_total: getVal(r, iIdx, 'farra_total'),
      s1_wr: getVal(r, iIdx, 's1_wr'),
      s2_wr: getVal(r, iIdx, 's2_wr'),
      s3_wr: getVal(r, iIdx, 's3_wr')
    };
  });

  const outRows = [];

  pRows.forEach(r => {
    const playerName = safe(r[pIdx['player_name']]);
    if (!playerName) return;

    outRows.push([
      ('player_id' in pIdx) ? safe(r[pIdx['player_id']]) : playerName,
      playerName,
      ('display_name' in pIdx) ? safe(r[pIdx['display_name']]) : playerName,
      ('photo_url' in pIdx) ? safe(r[pIdx['photo_url']]) : '',
      statusMap[playerName]?.division_actual || '-',
      rankMap[playerName]?.rank_position || '',
      rankMap[playerName]?.score_last14 || '',
      rankMap[playerName]?.avg_last14 || '',
      rachasMap[playerName]?.racha_actual || '',
      rachasMap[playerName]?.mejor_racha_g || '',
      rachasMap[playerName]?.mejor_racha_p || '',
      rachasMap[playerName]?.seq_5 || '',
      rachasMap[playerName]?.seq_10 || '',
      indivMap[playerName]?.match_wr ?? '',
      indivMap[playerName]?.sets_wr ?? '',
      indivMap[playerName]?.games_wr ?? '',
      indivMap[playerName]?.wr_2s ?? '',
      indivMap[playerName]?.wr_3s ?? '',
      indivMap[playerName]?.hard_sets_wr ?? '',
      indivMap[playerName]?.easy_sets_wr ?? '',
      indivMap[playerName]?.long_m_wr ?? '',
      indivMap[playerName]?.short_m_wr ?? '',
      indivMap[playerName]?.cbs_total ?? '',
      indivMap[playerName]?.farra_total ?? '',
      indivMap[playerName]?.s1_wr ?? '',
      indivMap[playerName]?.s2_wr ?? '',
      indivMap[playerName]?.s3_wr ?? ''
    ]);
  });

  out.clearContents();
  out.getRange(1, 1, 1, 27).setValues([[
    'player_id','player_name','display_name','photo_url',
    'division_actual','rank_position','score_last14','avg_last14',
    'racha_actual','mejor_racha_g','mejor_racha_p','seq_5','seq_10',
    'match_wr','sets_wr','games_wr','wr_2s','wr_3s',
    'hard_sets_wr','easy_sets_wr','long_m_wr','short_m_wr',
    'cbs_total','farra_total','s1_wr','s2_wr','s3_wr'
  ]]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, 27).setValues(outRows);

    [14,15,16,17,18,19,20,21,22,25,26,27].forEach(col => {
      out.getRange(2, col, outRows.length, 1).setNumberFormat('0.0%');
    });

    [23,24].forEach(col => {
      out.getRange(2, col, outRows.length, 1).setNumberFormat('0');
    });
  }

  notifyUser(
    'app_players reconstruida\n' +
    'Filas generadas: ' + outRows.length
  );
}
function getVal(row, idx, key) {
  return (key in idx) ? row[idx[key]] : '';
}
function buildCurrentPlayerStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cfg = ss.getSheetByName('current_league_config');
  const out = ss.getSheetByName('current_player_status');
  const playersSheet = ss.getSheetByName('players');
  const leaguePlayersSheet = ss.getSheetByName('league_players');

  if (!cfg) throw new Error('No existe la hoja current_league_config');
  if (!out) throw new Error('No existe la hoja current_player_status');
  if (!playersSheet) throw new Error('No existe la hoja players');
  if (!leaguePlayersSheet) throw new Error('No existe la hoja league_players');

  const cfgData = cfg.getDataRange().getValues();
if (cfgData.length < 2) throw new Error('current_league_config no tiene datos');

const cfgHeaders = cfgData[0].map(h => String(h).trim());
const cfgRows = cfgData.slice(1);

const cfgIdx = {};
cfgHeaders.forEach((h, i) => cfgIdx[h] = i);

if (!('setting' in cfgIdx) || !('value' in cfgIdx)) {
  throw new Error('current_league_config debe tener columnas setting y value');
}

let currentLeagueId = '';

cfgRows.forEach(r => {
  const setting = safe(r[cfgIdx['setting']]);
  const value = safe(r[cfgIdx['value']]);

  if (setting === 'current_league' || setting === 'current_league_id') {
    currentLeagueId = value;
  }
});

if (!currentLeagueId) throw new Error('No se encontró current_league en current_league_config');

  const pData = playersSheet.getDataRange().getValues();
  if (pData.length < 2) throw new Error('players no tiene datos');

  const pHeaders = pData[0].map(h => String(h).trim());
  const pRows = pData.slice(1);

  const pIdx = {};
  pHeaders.forEach((h, i) => pIdx[h] = i);

  if (!('player_name' in pIdx)) {
    throw new Error('players debe tener columna player_name');
  }

  const lpData = leaguePlayersSheet.getDataRange().getValues();
  if (lpData.length < 2) throw new Error('league_players no tiene datos');

  const lpHeaders = lpData[0].map(h => String(h).trim());
  const lpRows = lpData.slice(1);

  const lpIdx = {};
  lpHeaders.forEach((h, i) => lpIdx[h] = i);

  const requiredLp = [
    'league_id',
    'player_name',
    'registered_division',
    'role',
    'status',
    'replacement_division'
  ];

  const missingLp = requiredLp.filter(h => !(h in lpIdx));
  if (missingLp.length > 0) {
    throw new Error('Faltan columnas en league_players: ' + missingLp.join(', '));
  }

  const allPlayers = [];

  pRows.forEach(r => {
    const player = safe(r[pIdx['player_name']]);
    if (player) allPlayers.push(player);
  });

  const currentMap = {};

  lpRows.forEach(r => {
    const leagueId = safe(r[lpIdx['league_id']]);
    if (leagueId !== currentLeagueId) return;

    const player = safe(r[lpIdx['player_name']]);
    if (!player) return;

    const role = safe(r[lpIdx['role']]);
    const status = safe(r[lpIdx['status']]);
    const registeredDivision = safe(r[lpIdx['registered_division']]).toUpperCase();
    const replacementDivision = safe(r[lpIdx['replacement_division']]).toUpperCase();

    if (status && status !== 'Activo') return;

    if (role === 'Reemplazo') {
      currentMap[player] = 'Reemplazo';
    } else {
      currentMap[player] = registeredDivision || replacementDivision || '-';
    }

    if (!allPlayers.includes(player)) {
      allPlayers.push(player);
    }
  });

  const uniquePlayers = [...new Set(allPlayers)];

  const outRows = uniquePlayers
    .sort((a, b) => a.localeCompare(b))
    .map(player => {
      const isActive = player in currentMap ? 1 : 0;

      return [
        player,
        currentLeagueId,
        isActive ? currentMap[player] : '-',
        isActive
      ];
    });

  out.clearContents();

  out.getRange(1, 1, 1, 4).setValues([[
    'player',
    'current_league_id',
    'division_actual',
    'is_active_current_league'
  ]]);

  if (outRows.length > 0) {
    out.getRange(2, 1, outRows.length, 4).setValues(outRows);
  }

  notifyUser(
    'current_player_status reconstruida desde league_players\n' +
    'Liga actual: ' + currentLeagueId + '\n' +
    'Jugadores procesados: ' + outRows.length
  );
}


function notifyUser(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}
function initializePlayedOrderHistorical() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('matches_raw');

  if (!sh) throw new Error('No existe la hoja matches_raw');

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0].map(h => String(h).trim());
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const required = [
    'winner',
    'set1_p1',
    'set1_p2',
    'set2_p1',
    'set2_p2',
    'played_order',
    'played_at'
  ];

  const missing = required.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan columnas en matches_raw: ' + missing.join(', '));
  }

  const orderCol = idx['played_order'] + 1;
  const playedAtCol = idx['played_at'] + 1;

  let order = 0;
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const winner = String(row[idx['winner']] || '').trim();
    const s1p1 = row[idx['set1_p1']];
    const s1p2 = row[idx['set1_p2']];
    const s2p1 = row[idx['set2_p1']];
    const s2p2 = row[idx['set2_p2']];
    const currentOrder = row[idx['played_order']];

    const isPlayed =
      (winner === 'P1' || winner === 'P2') &&
      s1p1 !== '' && s1p1 !== null &&
      s1p2 !== '' && s1p2 !== null &&
      s2p1 !== '' && s2p1 !== null &&
      s2p2 !== '' && s2p2 !== null;

    if (isPlayed) {
      if (!currentOrder) {
        order++;
        const sheetRow = i + 1;
        sh.getRange(sheetRow, orderCol).setValue(order);

        if (!row[idx['played_at']]) {
          sh.getRange(sheetRow, playedAtCol).setValue(now);
        }
      } else {
        order = Math.max(order, Number(currentOrder) || 0);
      }
    }
  }

  notifyUser('played_order histórico inicializado correctamente');
}
function assignPlayedOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('matches_raw');

  if (!sh) throw new Error('No existe la hoja matches_raw');

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0].map(h => String(h).trim());
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const required = [
    'winner',
    'set1_p1',
    'set1_p2',
    'set2_p1',
    'set2_p2',
    'result_type',
    'played_order',
    'played_at'
  ];

  const missing = required.filter(h => !(h in idx));
  if (missing.length > 0) {
    throw new Error('Faltan columnas en matches_raw: ' + missing.join(', '));
  }

  let maxOrder = 0;

  for (let i = 1; i < data.length; i++) {
    const v = Number(data[i][idx['played_order']] || 0);
    if (v > maxOrder) maxOrder = v;
  }

  const orderCol = idx['played_order'] + 1;
  const playedAtCol = idx['played_at'] + 1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const winner = String(row[idx['winner']] || '').trim().toUpperCase();
    const resultType = String(row[idx['result_type']] || '').trim().toUpperCase();

    const s1p1 = row[idx['set1_p1']];
    const s1p2 = row[idx['set1_p2']];
    const s2p1 = row[idx['set2_p1']];
    const s2p2 = row[idx['set2_p2']];
    const currentOrder = row[idx['played_order']];

    const hasWinner = winner === 'P1' || winner === 'P2';

    const hasRealScore =
      s1p1 !== '' && s1p1 !== null &&
      s1p2 !== '' && s1p2 !== null &&
      s2p1 !== '' && s2p1 !== null &&
      s2p2 !== '' && s2p2 !== null;

    const isWO =
      resultType === 'WO_2_0' ||
      resultType === 'WO_INJURY_2_0' ||
      resultType === 'WO_INJURY_2_1';

    const isClosedMatch = hasWinner && (hasRealScore || isWO);

    if (isClosedMatch && !currentOrder) {
      maxOrder++;

      const sheetRow = i + 1;
      sh.getRange(sheetRow, orderCol).setValue(maxOrder);
      sh.getRange(sheetRow, playedAtCol).setValue(new Date());
    }
  }
}
function sortRowsByPlayedOrder(rows, idx) {
  if (!('played_order' in idx)) {
    throw new Error('Falta la columna played_order en matches_raw');
  }

  return rows
    .map((row, originalIndex) => {
      const rawOrder = row[idx['played_order']];
      const order = Number(rawOrder || 0);

      return {
        row,
        originalIndex,
        order: order > 0 ? order : Number.MAX_SAFE_INTEGER
      };
    })
    .sort((a, b) => {
      if (a.order !== b.order) return a.order - b.order;
      return a.originalIndex - b.originalIndex;
    })
    .map(item => item.row);
}
function syncReplacementPlayersIntoPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('matches_raw');
  const players = ss.getSheetByName('players');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!players) throw new Error('No existe la hoja players');

  const rawData = raw.getDataRange().getValues();
  const pData = players.getDataRange().getValues();

  if (rawData.length < 2) return;
  if (pData.length < 1) throw new Error('players no tiene encabezados');

  const rawHeaders = rawData[0].map(h => String(h).trim());
  const rawRows = rawData.slice(1);

  const pHeaders = pData[0].map(h => String(h).trim());
  const pRows = pData.slice(1);

  const rawIdx = {};
  rawHeaders.forEach((h, i) => rawIdx[h] = i);

  const pIdx = {};
  pHeaders.forEach((h, i) => pIdx[h] = i);

  const requiredRaw = [
    'pareja1_j1',
    'pareja1_j2',
    'pareja2_j1',
    'pareja2_j2'
  ];

  const missingRaw = requiredRaw.filter(h => !(h in rawIdx));
  if (missingRaw.length > 0) {
    throw new Error('Faltan columnas en matches_raw: ' + missingRaw.join(', '));
  }

  const requiredPlayers = [
    'player_id',
    'player_name',
    'display_name',
    'photo_url',
    'active',
    'division_actual',
    'notes'
  ];

  const missingPlayers = requiredPlayers.filter(h => !(h in pIdx));
  if (missingPlayers.length > 0) {
    throw new Error('Faltan columnas en players: ' + missingPlayers.join(', '));
  }

  const existingPlayers = {};

  pRows.forEach(r => {
    const name = safe(r[pIdx['player_name']]);
    if (!name) return;
    existingPlayers[name] = true;
  });

  const detectedPlayers = {};

  rawRows.forEach(r => {
    [
      safe(r[rawIdx['pareja1_j1']]),
      safe(r[rawIdx['pareja1_j2']]),
      safe(r[rawIdx['pareja2_j1']]),
      safe(r[rawIdx['pareja2_j2']])
    ].forEach(player => {
      if (!player) return;
      detectedPlayers[player] = true;
    });
  });

  const newRows = [];

  Object.keys(detectedPlayers)
    .sort((a, b) => a.localeCompare(b))
    .forEach(player => {
      if (existingPlayers[player]) return;

      const row = new Array(pHeaders.length).fill('');

      row[pIdx['player_id']] = makePlayerId(player);
      row[pIdx['player_name']] = player;
      row[pIdx['display_name']] = player;
      row[pIdx['photo_url']] = '';
      row[pIdx['active']] = true;
      row[pIdx['division_actual']] = '';
row[pIdx['notes']] = 'auto_from_match';

      newRows.push(row);
    });

  if (newRows.length > 0) {
    players
      .getRange(players.getLastRow() + 1, 1, newRows.length, pHeaders.length)
      .setValues(newRows);
  }

  Logger.log('Reemplazos creados automáticamente: ' + newRows.length);
}
function makePlayerId(playerName) {
  return String(playerName || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}
function syncLeaguePlayersFromMatches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const raw = ss.getSheetByName('matches_raw');
  const leaguePlayers = ss.getSheetByName('league_players');
  const leagues = ss.getSheetByName('leagues');

  if (!raw) throw new Error('No existe la hoja matches_raw');
  if (!leaguePlayers) throw new Error('No existe la hoja league_players');
  if (!leagues) throw new Error('No existe la hoja leagues');

  const rawData = raw.getDataRange().getValues();
  const lpData = leaguePlayers.getDataRange().getValues();
  const leaguesData = leagues.getDataRange().getValues();

  if (rawData.length < 2) return;
  if (lpData.length < 1) throw new Error('league_players no tiene encabezados');
  if (leaguesData.length < 2) throw new Error('leagues no tiene datos');

  const rawHeaders = rawData[0].map(h => String(h).trim());
  const rawRows = rawData.slice(1);

  const lpHeaders = lpData[0].map(h => String(h).trim());
  const lpRows = lpData.slice(1);

  const leaguesHeaders = leaguesData[0].map(h => String(h).trim());
  const leaguesRows = leaguesData.slice(1);

  const rawIdx = {};
  rawHeaders.forEach((h, i) => rawIdx[h] = i);

  const lpIdx = {};
  lpHeaders.forEach((h, i) => lpIdx[h] = i);

  const leaguesIdx = {};
  leaguesHeaders.forEach((h, i) => leaguesIdx[h] = i);

  const requiredRaw = [
    'campeonato',
    'division',
    'pareja1_j1',
    'pareja1_j2',
    'pareja2_j1',
    'pareja2_j2'
  ];

  const missingRaw = requiredRaw.filter(h => !(h in rawIdx));
  if (missingRaw.length > 0) {
    throw new Error('Faltan columnas en matches_raw: ' + missingRaw.join(', '));
  }

  const requiredLeaguePlayers = [
    'league_id',
    'client_id',
    'player_name',
    'registered_division',
    'role',
    'status',
    'source',
    'created_at',
    'replacement_division'
  ];

  const missingLeaguePlayers = requiredLeaguePlayers.filter(h => !(h in lpIdx));
  if (missingLeaguePlayers.length > 0) {
    throw new Error('Faltan columnas en league_players: ' + missingLeaguePlayers.join(', '));
  }

  const requiredLeagues = ['league_id', 'client_id'];
  const missingLeagues = requiredLeagues.filter(h => !(h in leaguesIdx));
  if (missingLeagues.length > 0) {
    throw new Error('Faltan columnas en leagues: ' + missingLeagues.join(', '));
  }

  const clientByLeague = {};

  leaguesRows.forEach(r => {
    const leagueId = safe(r[leaguesIdx['league_id']]);
    const clientId = safe(r[leaguesIdx['client_id']]);
    if (!leagueId) return;
    clientByLeague[leagueId] = clientId || '';
  });

  const existingLeaguePlayers = {};

  lpRows.forEach((r, i) => {
    const leagueId = safe(r[lpIdx['league_id']]);
    const player = safe(r[lpIdx['player_name']]);
    const role = safe(r[lpIdx['role']]);
    if (!leagueId || !player) return;

    existingLeaguePlayers[`${leagueId}|${player}`] = {
      rowNumber: i + 2,
      role
    };
  });

  const detectedLeaguePlayers = {};

  rawRows.forEach(r => {
    const leagueId = safe(r[rawIdx['campeonato']]);
    const division = safe(r[rawIdx['division']]);

    if (!leagueId) return;

    const names = [
      safe(r[rawIdx['pareja1_j1']]),
      safe(r[rawIdx['pareja1_j2']]),
      safe(r[rawIdx['pareja2_j1']]),
      safe(r[rawIdx['pareja2_j2']])
    ];

    names.forEach(player => {
      if (!player) return;

      const key = `${leagueId}|${player}`;

      if (!detectedLeaguePlayers[key]) {
        detectedLeaguePlayers[key] = {
          leagueId,
          player,
          divisions: {}
        };
      }

      if (division) {
        detectedLeaguePlayers[key].divisions[division] = true;
      }
    });
  });

  const newRows = [];
  const replacementDivisionCol = lpIdx['replacement_division'] + 1;

  Object.keys(detectedLeaguePlayers)
    .sort()
    .forEach(key => {
      const item = detectedLeaguePlayers[key];
      const leagueId = item.leagueId;
      const player = item.player;
      const clientId = clientByLeague[leagueId] || '';
      const replacementDivision = Object.keys(item.divisions).sort().join('/');

      const existing = existingLeaguePlayers[key];

      if (!existing) {
        const row = new Array(lpHeaders.length).fill('');

        row[lpIdx['league_id']] = leagueId;
        row[lpIdx['client_id']] = clientId;
        row[lpIdx['player_name']] = player;
        row[lpIdx['registered_division']] = '';
        row[lpIdx['role']] = 'Reemplazo';
        row[lpIdx['status']] = 'Activo';
        row[lpIdx['source']] = 'auto_from_match';
        row[lpIdx['created_at']] = new Date();
        row[lpIdx['replacement_division']] = replacementDivision;

        newRows.push(row);
        return;
      }

      if (existing.role === 'Reemplazo' && replacementDivision) {
        leaguePlayers
          .getRange(existing.rowNumber, replacementDivisionCol)
          .setValue(replacementDivision);
      }
    });

  if (newRows.length > 0) {
    leaguePlayers
      .getRange(leaguePlayers.getLastRow() + 1, 1, newRows.length, lpHeaders.length)
      .setValues(newRows);
  }

  Logger.log('syncLeaguePlayersFromMatches ejecutada. Nuevos reemplazos: ' + newRows.length);
}
function buildAppPlayerProfiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const appPlayersSheet = ss.getSheetByName('app_players');
const rankingSheet = ss.getSheetByName('ranking_last14');
const statsSheet = ss.getSheetByName('stats_indiv');
const rachasSheet = ss.getSheetByName('stats_rachas');
const standingsSheet = ss.getSheetByName('division_standings');
let out = ss.getSheetByName('app_player_profiles');

  if (!appPlayersSheet) throw new Error('No existe la hoja app_players');
if (!rankingSheet) throw new Error('No existe la hoja ranking_last14');
if (!statsSheet) throw new Error('No existe la hoja stats_indiv');
if (!standingsSheet) throw new Error('No existe la hoja division_standings');

  if (!out) {
    out = ss.insertSheet('app_player_profiles');
  }

  const appPlayers = sheetToObjects_(appPlayersSheet);
const ranking = sheetToObjects_(rankingSheet);
const stats = sheetToObjects_(statsSheet);
const rachas = rachasSheet ? sheetToObjects_(rachasSheet) : [];
const standings = sheetToObjects_(standingsSheet);

  const appByPlayer = indexByPlayer_(appPlayers, ['player_name', 'player']);
const rankingByPlayer = indexByPlayer_(ranking, ['player']);
const statsByPlayer = indexByPlayer_(stats, ['player']);
const rachasByPlayer = indexByPlayer_(rachas, ['player']);
const standingsByPlayerLeagueDiv = indexStandingsByPlayerLeagueDiv_(standings);

  const allPlayersSet = new Set();

  appPlayers.forEach(r => {
    const p = safeProfile_(r.player_name || r.player);
    if (p) allPlayersSet.add(p);
  });

  ranking.forEach(r => {
    const p = safeProfile_(r.player);
    if (p) allPlayersSet.add(p);
  });

  stats.forEach(r => {
    const p = safeProfile_(r.player);
    if (p) allPlayersSet.add(p);
  });

  const outputHeaders = [
  // Identidad
  'player',
  'photo_url',
  'initials',
  'role',

  // Estado actual
  'current_league_id',
  'division_actual',
  'division_table_position',
  'division_table_points',
  'division_table_record',
  'division_table_diff_games',
  'rank_position',
  'score_last14',
  'matches_played',
  'matches_counted',
  'record_last14',
  'result_mix_last14',
  'racha_actual',

  // Logros / badges
  'badge_1',
  'badge_1_type',
  'badge_2',
  'badge_2_type',
  'badge_3',
  'badge_3_type',
  'achievements_count',

  // Carrera LIG global
  'match_g',
  'match_p',
  'match_record',
  'match_wr',
  'sets_g',
  'sets_p',
  'sets_record',
  'sets_wr',
  'games_g',
  'games_p',
  'games_record',
  'games_wr',

  // Carrera LIG por formato
  'tot_2s',
  'pj_2s_pct',
  'pg_2s',
  'pp_2s',
  'record_2s',
  'wr_2s',
  'tot_3s',
  'pj_3s_pct',
  'pg_3s',
  'pp_3s',
  'record_3s',
  'wr_3s',

  // Rendimiento por división A
  'a_match_record',
  'a_match_wr',
  'a_sets_record',
  'a_sets_wr',
  'a_games_record',
  'a_games_wr',

  // Rendimiento por división B
  'b_match_record',
  'b_match_wr',
  'b_sets_record',
  'b_sets_wr',
  'b_games_record',
  'b_games_wr',

  // Rendimiento por división C
  'c_match_record',
  'c_match_wr',
  'c_sets_record',
  'c_sets_wr',
  'c_games_record',
  'c_games_wr',

  // Rendimiento por división D
  'd_match_record',
  'd_match_wr',
  'd_sets_record',
  'd_sets_wr',
  'd_games_record',
  'd_games_wr',

  // Perfil competitivo
  'cbs_total',
  'farra_total',
  'tb_total_g',
  'tb_total_p',
  'tb_record',
  'tb_total_wr',
  'tight_total_g',
  'tight_total_p',
  'tight_record',
  'tight_total_wr',
  'hard_sets_g',
  'hard_sets_p',
  'hard_sets_record',
  'hard_sets_wr',
  'easy_sets_g',
  'easy_sets_p',
  'easy_sets_record',
  'easy_sets_wr',
  's3_g',
  's3_p',
  's3_record',
  's3_wr',

  // Clutch / narrativa
  'clutch_index',
  'clutch_label',
  'clutch_summary',
  'main_tag',
  'secondary_tag',
  'profile_summary',

  // Rachas
  'best_w_streak',
  'best_l_streak',
  'last5_record',
  'last5_wr',
  'last10_record',
  'last10_wr',

  // Ranking y evolución
  'rank_snapshot_current',
  'rank_snapshot_previous',
  'rank_delta',
  'rank_delta_label',
  'snapshot_count',

  // Carrera por temporadas
  'seasons_played',
  'best_division',
  'best_finish',
  'last_season_played',
  'last_season_division',
  'last_season_finish',
  'last_season_result',

  // Historial / trazabilidad
  'total_matches_logged',
  'last_match_id',
  'last_match_date',
  'last_match_result',
  'last_match_score'
];

  const players = Array.from(allPlayersSet).sort((a, b) => a.localeCompare(b));

  const outRows = players.map(player => {
    const ap = appByPlayer[player] || {};
    const rk = rankingByPlayer[player] || {};
    const st = statsByPlayer[player] || {};
    const ra = rachasByPlayer[player] || {};
    const initials = initialsProfile_(player);

const role =
  valueProfile_(ap, 'role') ||
  valueProfile_(ap, 'player_role') ||
  '';

const badge1 = '';
const badge1Type = '';
const badge2 = '';
const badge2Type = '';
const badge3 = '';
const badge3Type = '';
const achievementsCount = '';

    const photoUrl = valueProfile_(ap, 'photo_url');
    const divisionActual =
      valueProfile_(ap, 'division_actual') ||
      valueProfile_(rk, 'division_actual') ||
      '-';

    const rankPosition =
      valueProfile_(rk, 'rank_position') ||
      valueProfile_(ap, 'rank_position') ||
      '';

const currentLeagueId = valueProfile_(rk, 'league_id') || 'Cocar26A';
const standingKey = `${player}|${currentLeagueId}|${divisionActual}`;
const standingRow = standingsByPlayerLeagueDiv[standingKey] || {};

const divisionTablePosition = valueProfile_(standingRow, 'pos');
const divisionTablePoints = valueProfile_(standingRow, 'pts');
const divisionTableRecord = `${valueProfile_(standingRow, 'pg') || 0}-${valueProfile_(standingRow, 'pp') || 0}`;
const divisionTableDiffGames = valueProfile_(standingRow, 'diff_games');

    const rachaActual =
      valueProfile_(ap, 'racha_actual') ||
      valueProfile_(ra, 'racha_actual') ||
      valueProfile_(ra, 'current_streak') ||
      '';

    const scoreLast14 = valueProfile_(rk, 'score_last14');
    const matchesPlayed = valueProfile_(rk, 'matches_played');
    const matchesCounted = valueProfile_(rk, 'matches_counted');
    const recordLast14 = valueProfile_(rk, 'record_last14');
    const resultMixLast14 = valueProfile_(rk, 'result_mix_last14');

    const matchG = numProfile_(st.match_g);
    const matchP = numProfile_(st.match_p);

    const setsG = numProfile_(st.sets_g);
    const setsP = numProfile_(st.sets_p);

    const gamesG = numProfile_(st.games_g);
    const gamesP = numProfile_(st.games_p);

    const pg2s = numProfile_(st.pg_2s);
    const pp2s = numProfile_(st.pp_2s);
    const pg3s = numProfile_(st.pg_3s);
    const pp3s = numProfile_(st.pp_3s);

    const tbG = numProfile_(st.tb_total_g);
    const tbP = numProfile_(st.tb_total_p);

    const tightG = numProfile_(st.tight_total_g);
    const tightP = numProfile_(st.tight_total_p);

    const hardG = numProfile_(st.hard_sets_g);
    const hardP = numProfile_(st.hard_sets_p);

    const easyG = numProfile_(st.easy_sets_g);
    const easyP = numProfile_(st.easy_sets_p);

    const longG = numProfile_(st.long_m_g);
    const longP = numProfile_(st.long_m_p);

    const shortG = numProfile_(st.short_m_g);
    const shortP = numProfile_(st.short_m_p);

    const aMatchRecord = recordText_(numProfile_(st.a_match_g), numProfile_(st.a_match_p));
const aSetsRecord = recordText_(numProfile_(st.a_sets_g), numProfile_(st.a_sets_p));
const aGamesRecord = recordText_(numProfile_(st.a_games_g), numProfile_(st.a_games_p));

const bMatchRecord = recordText_(numProfile_(st.b_match_g), numProfile_(st.b_match_p));
const bSetsRecord = recordText_(numProfile_(st.b_sets_g), numProfile_(st.b_sets_p));
const bGamesRecord = recordText_(numProfile_(st.b_games_g), numProfile_(st.b_games_p));

const cMatchRecord = recordText_(numProfile_(st.c_match_g), numProfile_(st.c_match_p));
const cSetsRecord = recordText_(numProfile_(st.c_sets_g), numProfile_(st.c_sets_p));
const cGamesRecord = recordText_(numProfile_(st.c_games_g), numProfile_(st.c_games_p));

const dMatchRecord = recordText_(numProfile_(st.d_match_g), numProfile_(st.d_match_p));
const dSetsRecord = recordText_(numProfile_(st.d_sets_g), numProfile_(st.d_sets_p));
const dGamesRecord = recordText_(numProfile_(st.d_games_g), numProfile_(st.d_games_p));

const s3G = numProfile_(st.s3_g);
const s3P = numProfile_(st.s3_p);

    const mainTag = buildMainPlayerTag_(st);
    const secondaryTag = buildSecondaryPlayerTag_(st);
    const profileSummary = buildPlayerProfileSummary_(player, st, rk);
    const clutchIndex = '';
const clutchLabel = '';
const clutchSummary = '';

const rankSnapshotCurrent = '';
const rankSnapshotPrevious = '';
const rankDelta = '';
const rankDeltaLabel = '';
const snapshotCount = '';

const seasonsPlayed = '';
const bestDivision = '';
const bestFinish = '';
const lastSeasonPlayed = '';
const lastSeasonDivision = '';
const lastSeasonFinish = '';
const lastSeasonResult = '';

const totalMatchesLogged = '';
const lastMatchId = '';
const lastMatchDate = '';
const lastMatchResult = '';
const lastMatchScore = '';

    return [
  // Identidad
  player,
  photoUrl,
  initials,
  role,

  // Estado actual
  currentLeagueId,
  divisionActual,
  divisionTablePosition,
  divisionTablePoints,
  divisionTableRecord,
  divisionTableDiffGames,
  rankPosition,
  scoreLast14,
  matchesPlayed,
  matchesCounted,
  recordLast14,
  resultMixLast14,
  rachaActual,

  // Logros / badges
  badge1,
  badge1Type,
  badge2,
  badge2Type,
  badge3,
  badge3Type,
  achievementsCount,

  // Carrera LIG global
  matchG,
  matchP,
  recordText_(matchG, matchP),
  valueProfile_(st, 'match_wr'),
  setsG,
  setsP,
  recordText_(setsG, setsP),
  valueProfile_(st, 'sets_wr'),
  gamesG,
  gamesP,
  recordText_(gamesG, gamesP),
  valueProfile_(st, 'games_wr'),

  // Carrera LIG por formato
  valueProfile_(st, 'tot_2s'),
  valueProfile_(st, 'pj_2s_pct'),
  pg2s,
  pp2s,
  recordText_(pg2s, pp2s),
  valueProfile_(st, 'wr_2s'),
  valueProfile_(st, 'tot_3s'),
  valueProfile_(st, 'pj_3s_pct'),
  pg3s,
  pp3s,
  recordText_(pg3s, pp3s),
  valueProfile_(st, 'wr_3s'),

  // Rendimiento por división A
  aMatchRecord,
  valueProfile_(st, 'a_match_wr'),
  aSetsRecord,
  valueProfile_(st, 'a_sets_wr'),
  aGamesRecord,
  valueProfile_(st, 'a_games_wr'),

  // Rendimiento por división B
  bMatchRecord,
  valueProfile_(st, 'b_match_wr'),
  bSetsRecord,
  valueProfile_(st, 'b_sets_wr'),
  bGamesRecord,
  valueProfile_(st, 'b_games_wr'),

  // Rendimiento por división C
  cMatchRecord,
  valueProfile_(st, 'c_match_wr'),
  cSetsRecord,
  valueProfile_(st, 'c_sets_wr'),
  cGamesRecord,
  valueProfile_(st, 'c_games_wr'),

  // Rendimiento por división D
  dMatchRecord,
  valueProfile_(st, 'd_match_wr'),
  dSetsRecord,
  valueProfile_(st, 'd_sets_wr'),
  dGamesRecord,
  valueProfile_(st, 'd_games_wr'),

  // Perfil competitivo
  valueProfile_(st, 'cbs_total'),
  valueProfile_(st, 'farra_total'),
  tbG,
  tbP,
  recordText_(tbG, tbP),
  valueProfile_(st, 'tb_total_wr'),
  tightG,
  tightP,
  recordText_(tightG, tightP),
  valueProfile_(st, 'tight_total_wr'),
  hardG,
  hardP,
  recordText_(hardG, hardP),
  valueProfile_(st, 'hard_sets_wr'),
  easyG,
  easyP,
  recordText_(easyG, easyP),
  valueProfile_(st, 'easy_sets_wr'),
  s3G,
  s3P,
  recordText_(s3G, s3P),
  valueProfile_(st, 's3_wr'),

  // Clutch / narrativa
  clutchIndex,
  clutchLabel,
  clutchSummary,
  mainTag,
  secondaryTag,
  profileSummary,

  // Rachas
  valueProfile_(ra, 'best_w_streak') || valueProfile_(ra, 'bestW') || '',
  valueProfile_(ra, 'best_l_streak') || valueProfile_(ra, 'bestL') || '',
  valueProfile_(ra, 'last5_record') || '',
  valueProfile_(ra, 'last5_wr') || '',
  valueProfile_(ra, 'last10_record') || '',
  valueProfile_(ra, 'last10_wr') || '',

  // Ranking y evolución
  rankSnapshotCurrent,
  rankSnapshotPrevious,
  rankDelta,
  rankDeltaLabel,
  snapshotCount,

  // Carrera por temporadas
  seasonsPlayed,
  bestDivision,
  bestFinish,
  lastSeasonPlayed,
  lastSeasonDivision,
  lastSeasonFinish,
  lastSeasonResult,

  // Historial / trazabilidad
  totalMatchesLogged,
  lastMatchId,
  lastMatchDate,
  lastMatchResult,
  lastMatchScore
];
  });

  out.clear();

out.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

// Deja toda la hoja en formato texto para evitar que Google Sheets convierta
// registros tipo 7-4, 1-10 o 5-2-1-3 en fechas.
out.getRange(1, 1, out.getMaxRows(), outputHeaders.length).setNumberFormat('@');

if (outRows.length > 0) {
  out.getRange(2, 1, outRows.length, outputHeaders.length).setValues(outRows);
}

out.autoResizeColumns(1, outputHeaders.length);

  notifyUser(
    'app_player_profiles reconstruida\n' +
    'Jugadores procesados: ' + outRows.length
  );
}

function sheetToObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h || '').trim());
  const rows = data.slice(1);

  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (h) obj[h] = row[i];
    });
    return obj;
  });
}

function indexByPlayer_(rows, possibleKeys) {
  const map = {};

  rows.forEach(row => {
    let player = '';

    for (const key of possibleKeys) {
      if (row[key]) {
        player = safeProfile_(row[key]);
        break;
      }
    }

    if (player) map[player] = row;
  });

  return map;
}

function safeProfile_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function valueProfile_(obj, key) {
  if (!obj || !(key in obj)) return '';
  return obj[key] === null || obj[key] === undefined ? '' : obj[key];
}

function numProfile_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function initialsProfile_(name) {
  const parts = String(name || '')
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();

  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function recordText_(g, p) {
  return `${Number(g || 0)}-${Number(p || 0)}`;
}

function pctNumberProfile_(v) {
  if (v === '' || v === null || v === undefined) return 0;

  if (typeof v === 'number') {
    return v <= 1 ? v * 100 : v;
  }

  const s = String(v).replace('%', '').trim();
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function buildMainPlayerTag_(st) {
  const hardWr = pctNumberProfile_(st.hard_sets_wr);
  const easyWr = pctNumberProfile_(st.easy_sets_wr);
  const wr3s = pctNumberProfile_(st.wr_3s);
  const wr2s = pctNumberProfile_(st.wr_2s);
  const tbWr = pctNumberProfile_(st.tb_total_wr);

  const hardPlayed = numProfile_(st.hard_sets_g) + numProfile_(st.hard_sets_p);
  const easyPlayed = numProfile_(st.easy_sets_g) + numProfile_(st.easy_sets_p);
  const tbPlayed = numProfile_(st.tb_total_g) + numProfile_(st.tb_total_p);

  if (tbPlayed >= 3 && tbWr >= 70) return 'Tie break killer';
  if (hardPlayed >= 6 && hardWr >= 65) return 'Competidor de sets duros';
  if (wr3s >= 65 && (numProfile_(st.pg_3s) + numProfile_(st.pp_3s)) >= 3) return 'Jugador de partidos largos';
  if (wr2s >= 70 && (numProfile_(st.pg_2s) + numProfile_(st.pp_2s)) >= 4) return 'Dominante en partidos cortos';
  if (easyPlayed >= 8 && easyWr >= 70) return 'Dominante cuando abre ventaja';

  return 'Perfil en construcción';
}

function buildSecondaryPlayerTag_(st) {
  const cbs = numProfile_(st.cbs_total);
  const farra = numProfile_(st.farra_total);
  const matchWr = pctNumberProfile_(st.match_wr);
  const gamesWr = pctNumberProfile_(st.games_wr);

  if (cbs >= 2 && cbs > farra) return 'Alta capacidad de remontada';
  if (farra >= 2 && farra > cbs) return 'Debe cerrar mejor los partidos';
  if (matchWr >= 60 && gamesWr < 52) return 'Gana partidos muy ajustados';
  if (matchWr < 45 && gamesWr >= 50) return 'Compite más de lo que gana';

  return '';
}

function buildPlayerProfileSummary_(player, st, rk) {
  const matchRecord = recordText_(numProfile_(st.match_g), numProfile_(st.match_p));
  const matchWr = valueProfile_(st, 'match_wr') || '0%';
  const recordL14 = valueProfile_(rk, 'record_last14') || '0-0';

  return `${player} registra ${matchRecord} en partidos (${matchWr}) y un Last14 de ${recordL14}.`;
}
function indexStandingsByPlayerLeagueDiv_(standingsRows) {
  const map = {};

  standingsRows.forEach(row => {
    const player = safeProfile_(row.player);
    const leagueId = safeProfile_(row.league_id);
    const division = safeProfile_(row.division);

    if (!player || !leagueId || !division) return;

    const key = `${player}|${leagueId}|${division}`;

    map[key] = {
      pos: keepZero_(row.pos),
      pts: keepZero_(row.pts),
      pj: keepZero_(row.pj),
      pg: keepZero_(row.pg),
      pp: keepZero_(row.pp),
      diff_games: keepZero_(row.diff_games)
    };
  });

  return map;
}

function keepZero_(v) {
  return v === null || v === undefined ? '' : v;
}