(() => {
  const COLS = 10;
  const ROWS = 20;
  const BLOCK = 30;

  const boardCanvas = document.getElementById('board');
  const boardCtx = boardCanvas.getContext('2d');
  const nextCanvas = document.getElementById('next');
  const nextCtx = nextCanvas.getContext('2d');
  const scoreEl = document.getElementById('score');
  const levelEl = document.getElementById('level');
  const linesEl = document.getElementById('lines');
  const startBtn = document.getElementById('start-btn');
  const overlay = document.getElementById('overlay');
  const overlayText = document.getElementById('overlay-text');

  const COLORS = {
    I: '#00f0f0',
    O: '#f0f000',
    T: '#a000f0',
    S: '#00f000',
    Z: '#f00000',
    J: '#0000f0',
    L: '#f0a000',
  };

  const SHAPES = {
    I: [
      [0, 1], [1, 1], [2, 1], [3, 1],
    ],
    O: [
      [1, 0], [2, 0], [1, 1], [2, 1],
    ],
    T: [
      [1, 0], [0, 1], [1, 1], [2, 1],
    ],
    S: [
      [1, 0], [2, 0], [0, 1], [1, 1],
    ],
    Z: [
      [0, 0], [1, 0], [1, 1], [2, 1],
    ],
    J: [
      [0, 0], [0, 1], [1, 1], [2, 1],
    ],
    L: [
      [2, 0], [0, 1], [1, 1], [2, 1],
    ],
  };

  const PIECE_TYPES = Object.keys(SHAPES);

  function createEmptyBoard() {
    return Array.from({ length: ROWS }, () => Array(COLS).fill(null));
  }

  let board = createEmptyBoard();
  let current = null;
  let next = null;
  let score = 0;
  let level = 1;
  let lines = 0;
  let dropInterval = 1000;
  let dropTimer = null;
  let running = false;
  let paused = false;
  let gameOver = false;

  function randomPiece() {
    const type = PIECE_TYPES[Math.floor(Math.random() * PIECE_TYPES.length)];
    return {
      type,
      cells: SHAPES[type].map(([x, y]) => [x, y]),
      x: 3,
      y: -1,
      rotation: 0,
    };
  }

  function rotateCells(cells, type) {
    if (type === 'O') return cells.map(([x, y]) => [x, y]);
    // rotate around a pivot using a 4x4 box (SRS-like, simplified)
    const size = type === 'I' ? 4 : 3;
    return cells.map(([x, y]) => {
      const nx = size - 1 - y;
      const ny = x;
      return [nx, ny];
    });
  }

  function getCellCoords(piece) {
    return piece.cells.map(([x, y]) => [piece.x + x, piece.y + y]);
  }

  function collides(piece, offsetX = 0, offsetY = 0, cells = null) {
    const testCells = cells || piece.cells;
    for (const [cx, cy] of testCells) {
      const x = piece.x + cx + offsetX;
      const y = piece.y + cy + offsetY;
      if (x < 0 || x >= COLS || y >= ROWS) return true;
      if (y >= 0 && board[y][x]) return true;
    }
    return false;
  }

  function mergePiece() {
    for (const [x, y] of getCellCoords(current)) {
      if (y >= 0) board[y][x] = current.type;
    }
  }

  function clearLines() {
    let cleared = 0;
    for (let y = ROWS - 1; y >= 0; y--) {
      if (board[y].every((cell) => cell !== null)) {
        board.splice(y, 1);
        board.unshift(Array(COLS).fill(null));
        cleared++;
        y++;
      }
    }
    if (cleared > 0) {
      const points = [0, 100, 300, 500, 800][cleared] * level;
      score += points;
      lines += cleared;
      level = Math.floor(lines / 10) + 1;
      dropInterval = Math.max(100, 1000 - (level - 1) * 80);
      updateStats();
    }
  }

  function updateStats() {
    scoreEl.textContent = score;
    levelEl.textContent = level;
    linesEl.textContent = lines;
  }

  function spawnPiece() {
    current = next || randomPiece();
    next = randomPiece();
    current.x = 3;
    current.y = -1;
    drawNext();
    if (collides(current)) {
      endGame();
    }
  }

  function endGame() {
    gameOver = true;
    running = false;
    clearTimeout(dropTimer);
    overlayText.textContent = `ゲームオーバー\nスコア: ${score}`;
    overlay.classList.remove('hidden');
    startBtn.textContent = 'リスタート';
  }

  function move(dx) {
    if (!running || paused) return;
    if (!collides(current, dx, 0)) {
      current.x += dx;
      draw();
    }
  }

  function softDrop() {
    if (!running || paused) return;
    if (!collides(current, 0, 1)) {
      current.y += 1;
      score += 1;
      updateStats();
      draw();
    } else {
      lockPiece();
    }
  }

  function hardDrop() {
    if (!running || paused) return;
    let dropped = 0;
    while (!collides(current, 0, 1)) {
      current.y += 1;
      dropped++;
    }
    score += dropped * 2;
    updateStats();
    lockPiece();
  }

  function rotate() {
    if (!running || paused) return;
    const rotated = rotateCells(current.cells, current.type);
    const kicks = [0, -1, 1, -2, 2];
    for (const k of kicks) {
      if (!collides(current, k, 0, rotated)) {
        current.cells = rotated;
        current.x += k;
        draw();
        return;
      }
    }
  }

  function lockPiece() {
    mergePiece();
    clearLines();
    spawnPiece();
    draw();
  }

  function tick() {
    if (!running || paused) return;
    if (!collides(current, 0, 1)) {
      current.y += 1;
    } else {
      lockPiece();
    }
    draw();
    dropTimer = setTimeout(tick, dropInterval);
  }

  function drawCell(ctx, x, y, color) {
    ctx.fillStyle = color;
    ctx.fillRect(x * BLOCK, y * BLOCK, BLOCK, BLOCK);
    ctx.strokeStyle = 'rgba(0,0,0,0.3)';
    ctx.lineWidth = 1;
    ctx.strokeRect(x * BLOCK + 0.5, y * BLOCK + 0.5, BLOCK - 1, BLOCK - 1);
    ctx.fillStyle = 'rgba(255,255,255,0.2)';
    ctx.fillRect(x * BLOCK, y * BLOCK, BLOCK, 4);
  }

  function draw() {
    boardCtx.fillStyle = '#0f0f23';
    boardCtx.fillRect(0, 0, boardCanvas.width, boardCanvas.height);

    for (let y = 0; y < ROWS; y++) {
      for (let x = 0; x < COLS; x++) {
        if (board[y][x]) {
          drawCell(boardCtx, x, y, COLORS[board[y][x]]);
        }
      }
    }

    if (current) {
      for (const [x, y] of getCellCoords(current)) {
        if (y >= 0) drawCell(boardCtx, x, y, COLORS[current.type]);
      }
    }
  }

  function drawNext() {
    nextCtx.fillStyle = '#1a1a2e';
    nextCtx.fillRect(0, 0, nextCanvas.width, nextCanvas.height);
    if (!next) return;
    const cells = SHAPES[next.type];
    const size = next.type === 'I' ? 4 : 3;
    const scale = 24;
    const offsetX = (nextCanvas.width - size * scale) / 2;
    const offsetY = (nextCanvas.height - size * scale) / 2;
    for (const [x, y] of cells) {
      nextCtx.fillStyle = COLORS[next.type];
      nextCtx.fillRect(offsetX + x * scale, offsetY + y * scale, scale, scale);
      nextCtx.strokeStyle = 'rgba(0,0,0,0.3)';
      nextCtx.strokeRect(offsetX + x * scale + 0.5, offsetY + y * scale + 0.5, scale - 1, scale - 1);
    }
  }

  function togglePause() {
    if (!running || gameOver) return;
    paused = !paused;
    if (paused) {
      clearTimeout(dropTimer);
      overlayText.textContent = '一時停止中';
      overlay.classList.remove('hidden');
    } else {
      overlay.classList.add('hidden');
      dropTimer = setTimeout(tick, dropInterval);
    }
  }

  function startGame() {
    clearTimeout(dropTimer);
    board = createEmptyBoard();
    score = 0;
    level = 1;
    lines = 0;
    dropInterval = 1000;
    running = true;
    paused = false;
    gameOver = false;
    updateStats();
    overlay.classList.add('hidden');
    startBtn.textContent = 'リスタート';
    next = randomPiece();
    spawnPiece();
    draw();
    dropTimer = setTimeout(tick, dropInterval);
  }

  startBtn.addEventListener('click', startGame);

  document.addEventListener('keydown', (e) => {
    if (!running && e.code !== 'KeyP') return;
    switch (e.code) {
      case 'ArrowLeft':
        e.preventDefault();
        move(-1);
        break;
      case 'ArrowRight':
        e.preventDefault();
        move(1);
        break;
      case 'ArrowDown':
        e.preventDefault();
        softDrop();
        break;
      case 'ArrowUp':
        e.preventDefault();
        rotate();
        break;
      case 'Space':
        e.preventDefault();
        hardDrop();
        break;
      case 'KeyP':
        e.preventDefault();
        togglePause();
        break;
    }
  });

  draw();
  drawNext();
})();
