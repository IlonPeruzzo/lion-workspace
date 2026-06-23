#!/usr/bin/env python3
"""
Lion Workspace - Motion Tracker (native Python + OpenCV)

Trackeia 1+ pontos num video usando CSRT, KCF, MOSSE, MIL ou Lucas-Kanade.
Aplica smoothing (Whittaker / Kalman / Supersmoother).
Saida: JSON em stdout pra ser consumido pelo Electron.

Uso:
    python motion_tracker.py <video> --points "[[x,y],[x,y]]" --algorithm CSRT --smoothing whittaker

Saida (JSON):
    {
      "success": true,
      "algorithm": "CSRT",
      "metadata": {"fps": 30, "width": 1920, "height": 1080, "frame_count": 120, "duration": 4.0},
      "tracks": [
        {
          "id": 0,
          "history": [
            {"frame": 0, "time": 0.0, "x": 100.0, "y": 200.0, "scale": 1.0, "confidence": 1.0, "lost": false},
            ...
          ]
        },
        ...
      ]
    }
"""
import argparse
import json
import os
import sys
import traceback
import numpy as np
import cv2


# ====== ALGORITMOS DE TRACKING ======

def _create_tracker(name):
    """Cria tracker do OpenCV pelo nome. Retorna None se nao suportado."""
    name = (name or '').upper()
    # cv2.legacy.* pra OpenCV >= 4.5.1
    factories = {
        'CSRT':  getattr(cv2.legacy, 'TrackerCSRT_create', None),
        'KCF':   getattr(cv2.legacy, 'TrackerKCF_create', None),
        'MOSSE': getattr(cv2.legacy, 'TrackerMOSSE_create', None),
        'MIL':   getattr(cv2.legacy, 'TrackerMIL_create', None),
    }
    fn = factories.get(name)
    return fn() if fn else None


def _confidence(new_bbox, last_bbox, search_scale, init_size):
    """Confidence multi-fator: distancia + tamanho. 0.0 a 1.0."""
    if last_bbox is None or new_bbox is None:
        return 1.0
    lcx = last_bbox[0] + last_bbox[2] / 2
    lcy = last_bbox[1] + last_bbox[3] / 2
    ncx = new_bbox[0] + new_bbox[2] / 2
    ncy = new_bbox[1] + new_bbox[3] / 2
    dist = ((ncx - lcx) ** 2 + (ncy - lcy) ** 2) ** 0.5
    max_dist = max(init_size) * search_scale
    dist_conf = max(0.0, 1.0 - dist / max_dist) if max_dist > 0 else 1.0
    la = last_bbox[2] * last_bbox[3]
    na = new_bbox[2] * new_bbox[3]
    size_conf = min(la, na) / max(la, na) if max(la, na) > 0 else 0.0
    return (dist_conf + size_conf) / 2.0


def track_correlation_filter(cap, fps, width, height, point, opts):
    """Track usando CSRT/KCF/MOSSE/MIL — funciona com bbox que muda de escala."""
    algo = opts['algorithm']
    bbox_w = opts['bbox_width']
    bbox_h = opts['bbox_height']
    conf_threshold = opts['confidence_threshold']
    max_lost = opts['max_lost_frames']
    search_scale = opts['search_window_scale']

    tracker = _create_tracker(algo)
    if tracker is None:
        return None, f"Algoritmo {algo} nao disponivel no OpenCV instalado"

    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
    ret, frame = cap.read()
    if not ret:
        return None, "Nao consegui ler primeiro frame"

    sx, sy = int(point[0]), int(point[1])
    bbox = (
        max(0, sx - bbox_w // 2),
        max(0, sy - bbox_h // 2),
        min(bbox_w, width - sx + bbox_w // 2),
        min(bbox_h, height - sy + bbox_h // 2),
    )
    init_size = (bbox[2], bbox[3])
    init_area = bbox[2] * bbox[3]
    if not tracker.init(frame, bbox):
        return None, "Falha ao inicializar tracker"

    history = []
    history.append({
        'frame': 0, 'time': 0.0,
        'x': float(sx), 'y': float(sy),
        'scale': 1.0, 'confidence': 1.0, 'lost': False,
    })

    last_bbox = bbox
    consecutive_lost = 0
    frame_num = 0

    while True:
        ret, frame = cap.read()
        if not ret:
            break
        frame_num += 1
        t = frame_num / fps if fps > 0 else frame_num
        success, new_bbox = tracker.update(frame)
        confidence = _confidence(new_bbox, last_bbox, search_scale, init_size) if success else 0.0

        if success and confidence >= conf_threshold:
            cx = new_bbox[0] + new_bbox[2] / 2
            cy = new_bbox[1] + new_bbox[3] / 2
            scale = (new_bbox[2] * new_bbox[3] / init_area) ** 0.5 if init_area > 0 else 1.0
            last_bbox = new_bbox
            consecutive_lost = 0
            history.append({
                'frame': frame_num, 'time': t,
                'x': float(cx), 'y': float(cy),
                'scale': float(scale), 'confidence': float(confidence), 'lost': False,
            })
        else:
            consecutive_lost += 1
            history.append({
                'frame': frame_num, 'time': t,
                'x': None, 'y': None, 'scale': None,
                'confidence': 0.0, 'lost': True,
            })
            if consecutive_lost >= max_lost:
                # Continua adicionando "lost" pra completar timeline
                pass

    return history, None


def track_lucas_kanade(cap, fps, width, height, point, opts):
    """Track usando KLT optical flow com forward-backward validation."""
    conf_threshold = opts['confidence_threshold']
    max_lost = opts['max_lost_frames']
    win_size = opts.get('window_size', 21)
    pyramid_levels = opts.get('pyramid_levels', 5)
    max_iter = opts.get('max_iterations', 30)
    quality = opts.get('tracking_quality', 0.03)

    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
    ret, frame = cap.read()
    if not ret:
        return None, "Nao consegui ler primeiro frame"

    old_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    lk_params = dict(
        winSize=(win_size, win_size),
        maxLevel=pyramid_levels,
        criteria=(cv2.TERM_CRITERIA_EPS | cv2.TERM_CRITERIA_COUNT, max_iter, quality),
    )

    sx, sy = float(point[0]), float(point[1])
    p0 = np.array([[[sx, sy]]], dtype=np.float32)

    history = [{
        'frame': 0, 'time': 0.0,
        'x': sx, 'y': sy, 'scale': 1.0, 'confidence': 1.0, 'lost': False,
    }]

    consecutive_lost = 0
    frame_num = 0

    while True:
        ret, frame = cap.read()
        if not ret:
            break
        frame_num += 1
        t = frame_num / fps if fps > 0 else frame_num
        frame_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        try:
            p1, st, err = cv2.calcOpticalFlowPyrLK(old_gray, frame_gray, p0, None, **lk_params)
        except Exception:
            p1, st, err = None, None, None

        # Forward-backward validation
        bf_err = 0.0
        if p1 is not None and st is not None and len(st) > 0 and st[0][0] == 1:
            try:
                p0r, st_back, _ = cv2.calcOpticalFlowPyrLK(frame_gray, old_gray, p1, None, **lk_params)
                if p0r is not None and len(st_back) > 0 and st_back[0][0] == 1:
                    bf_err = ((p0[0][0][0] - p0r[0][0][0]) ** 2 + (p0[0][0][1] - p0r[0][0][1]) ** 2) ** 0.5
                else:
                    bf_err = 10.0
            except Exception:
                pass

        valid = (
            p1 is not None and st is not None and len(st) > 0
            and st[0][0] == 1 and bf_err < 2.0
        )

        if valid:
            new_x, new_y = float(p1[0][0][0]), float(p1[0][0][1])
            tracking_err = float(err[0][0]) if err is not None else 0.0
            err_conf = max(0.0, 1.0 - tracking_err / 100.0)
            # Confidence motion: penaliza saltos
            last = history[-1]
            if last['x'] is not None:
                motion = ((new_x - last['x']) ** 2 + (new_y - last['y']) ** 2) ** 0.5
                motion_conf = max(0.0, 1.0 - motion / (width * 0.1))
            else:
                motion_conf = 1.0
            margin = 50
            boundary_conf = 0.5 if (new_x < margin or new_x > width - margin or new_y < margin or new_y > height - margin) else 1.0
            confidence = err_conf * 0.5 + motion_conf * 0.4 + boundary_conf * 0.1

            if 0 <= new_x < width and 0 <= new_y < height and confidence >= conf_threshold:
                consecutive_lost = 0
                history.append({
                    'frame': frame_num, 'time': t,
                    'x': new_x, 'y': new_y, 'scale': 1.0,
                    'confidence': float(confidence), 'lost': False,
                })
                p0 = p1.copy()
                old_gray = frame_gray
                continue

        # Lost
        consecutive_lost += 1
        history.append({
            'frame': frame_num, 'time': t,
            'x': None, 'y': None, 'scale': None,
            'confidence': 0.0, 'lost': True,
        })
        old_gray = frame_gray
        if consecutive_lost > max_lost:
            # nao reinicializa, continua marcando como lost
            pass

    return history, None


# ====== SMOOTHING ======

def _whittaker_smooth(values, lam=100.0):
    """Whittaker-Eilers smoother — penaliza segunda derivada."""
    arr = np.asarray(values, dtype=np.float64)
    n = len(arr)
    if n < 3:
        return arr
    I = np.eye(n)
    D = np.diff(I, n=2, axis=0)
    A = I + lam * D.T @ D
    try:
        return np.linalg.solve(A, arr)
    except np.linalg.LinAlgError:
        return arr


def apply_whittaker(history, lam=100.0):
    """Aplica Whittaker em x, y do tracking. Pula pontos lost."""
    valid_idx = [i for i, p in enumerate(history) if not p.get('lost') and p.get('x') is not None]
    if len(valid_idx) < 3:
        return history
    xs = [history[i]['x'] for i in valid_idx]
    ys = [history[i]['y'] for i in valid_idx]
    sx = _whittaker_smooth(xs, lam)
    sy = _whittaker_smooth(ys, lam)
    for k, i in enumerate(valid_idx):
        history[i]['x'] = float(sx[k])
        history[i]['y'] = float(sy[k])
    # Scale smoothing tambem
    scales = [p.get('scale') for p in history if not p.get('lost') and p.get('scale') is not None]
    if len(scales) >= 3:
        ss = _whittaker_smooth(scales, lam)
        k = 0
        for p in history:
            if not p.get('lost') and p.get('scale') is not None:
                p['scale'] = float(ss[k])
                k += 1
    return history


def apply_kalman(history, process_noise=0.5, measurement_noise=1.0):
    """Kalman filter 4-state (x, y, vx, vy) com 1-3 passes baseado em noise."""
    valid_idx = [i for i, p in enumerate(history) if not p.get('lost') and p.get('x') is not None]
    if len(valid_idx) < 2:
        return history

    num_passes = 1
    if measurement_noise > 10.0:
        num_passes = 3
    elif measurement_noise > 5.0:
        num_passes = 2

    for _ in range(num_passes):
        kf = cv2.KalmanFilter(4, 2)
        damping = 1.0
        if measurement_noise > 20.0:
            damping = 0.5
        elif measurement_noise > 10.0:
            damping = 0.7
        kf.transitionMatrix = np.array([
            [1, 0, damping, 0],
            [0, 1, 0, damping],
            [0, 0, damping, 0],
            [0, 0, 0, damping],
        ], dtype=np.float32)
        kf.measurementMatrix = np.array([[1, 0, 0, 0], [0, 1, 0, 0]], dtype=np.float32)

        effective_pn = process_noise * (0.1 if measurement_noise > 50.0 else 1.0)
        q = np.eye(4, dtype=np.float32)
        q[2:4, 2:4] *= 0.1
        kf.processNoiseCov = effective_pn * q
        kf.measurementNoiseCov = measurement_noise * np.eye(2, dtype=np.float32)
        kf.errorCovPost = measurement_noise * np.eye(4, dtype=np.float32)

        first = history[valid_idx[0]]
        kf.statePre = np.array([first['x'], first['y'], 0, 0], dtype=np.float32)
        kf.statePost = np.array([first['x'], first['y'], 0, 0], dtype=np.float32)

        for i in valid_idx:
            kf.predict()
            measurement = np.array([history[i]['x'], history[i]['y']], dtype=np.float32)
            kf.correct(measurement)
            history[i]['x'] = float(kf.statePost[0])
            history[i]['y'] = float(kf.statePost[1])
    return history


def _tricube(u):
    u = np.abs(u)
    return np.where(u <= 1, (1 - u ** 3) ** 3, 0.0)


def _local_linear(t, y, span):
    n = len(t)
    out = np.zeros(n)
    for i in range(n):
        dist = np.abs(t - t[i])
        idx = np.argsort(dist)[:span]
        tl, yl = t[idx], y[idx]
        md = np.max(np.abs(tl - t[i]))
        w = _tricube(np.abs(tl - t[i]) / md) if md > 0 else np.ones(len(tl))
        try:
            W = np.diag(w)
            X = np.column_stack([np.ones(len(tl)), tl])
            c = np.linalg.solve(X.T @ W @ X, X.T @ W @ yl)
            out[i] = c[0] + c[1] * t[i]
        except np.linalg.LinAlgError:
            out[i] = np.average(yl, weights=w)
    return out


def _supersmoother_1d(t, y, alpha=0.3):
    n = len(t)
    spans = [max(5, int(0.05 * n)), max(10, int(0.2 * n)), max(20, int(0.5 * n))]
    cands = [_local_linear(t, y, s) for s in spans]
    # cross-validation residual (simplificado)
    res = []
    for s, c in zip(spans, cands):
        res.append(float(np.sum((y - c) ** 2)))
    best = int(np.argmin(res))
    out = np.zeros(n)
    min_res = min(res) if min(res) > 0 else 1.0
    for i in range(n):
        w = np.array([
            (1 - alpha) * np.exp(-res[j] / min_res) if j == best else alpha / 2.0
            for j in range(3)
        ])
        w /= np.sum(w)
        out[i] = sum(w[j] * cands[j][i] for j in range(3))
    return out


def apply_supersmoother(history, alpha=0.3):
    """Supersmoother (Friedman) — escolhe span adaptivo via cross-validation."""
    valid_idx = [i for i, p in enumerate(history) if not p.get('lost') and p.get('x') is not None]
    if len(valid_idx) < 5:
        return history
    t = np.array(valid_idx, dtype=np.float64)
    xs = np.array([history[i]['x'] for i in valid_idx], dtype=np.float64)
    ys = np.array([history[i]['y'] for i in valid_idx], dtype=np.float64)
    sx = _supersmoother_1d(t, xs, alpha)
    sy = _supersmoother_1d(t, ys, alpha)
    for k, i in enumerate(valid_idx):
        history[i]['x'] = float(sx[k])
        history[i]['y'] = float(sy[k])
    return history


# ====== ENTRY POINT ======

def main():
    p = argparse.ArgumentParser(description='Lion Workspace Motion Tracker')
    p.add_argument('video_path')
    p.add_argument('--points', required=True, help='JSON array de [x,y] iniciais')
    p.add_argument('--algorithm', default='CSRT', choices=['CSRT', 'KCF', 'MOSSE', 'MIL', 'LUCAS_KANADE'])
    p.add_argument('--smoothing', default='whittaker', choices=['whittaker', 'kalman', 'supersmoother', 'none'])
    p.add_argument('--bbox-width', type=int, default=40)
    p.add_argument('--bbox-height', type=int, default=40)
    p.add_argument('--confidence-threshold', type=float, default=0.3)
    p.add_argument('--max-lost-frames', type=int, default=5)
    p.add_argument('--search-window-scale', type=float, default=2.0)
    p.add_argument('--whittaker-lambda', type=float, default=100.0)
    p.add_argument('--kalman-process-noise', type=float, default=0.5)
    p.add_argument('--kalman-measurement-noise', type=float, default=1.0)
    p.add_argument('--supersmoother-alpha', type=float, default=0.3)
    p.add_argument('--window-size', type=int, default=21)
    p.add_argument('--pyramid-levels', type=int, default=5)
    p.add_argument('--max-iterations', type=int, default=30)
    p.add_argument('--tracking-quality', type=float, default=0.03)
    p.add_argument('--start-frame', type=int, default=0)
    p.add_argument('--end-frame', type=int, default=-1)
    args = p.parse_args()

    result = {
        'success': False,
        'algorithm': args.algorithm,
        'video_path': args.video_path,
        'metadata': {},
        'tracks': [],
        'errors': [],
    }

    try:
        if not os.path.exists(args.video_path):
            result['errors'].append(f"Arquivo nao encontrado: {args.video_path}")
            print(json.dumps(result))
            sys.exit(1)

        points = json.loads(args.points)
        if not isinstance(points, list) or not points:
            result['errors'].append("'--points' precisa ser um JSON array de [x,y]")
            print(json.dumps(result))
            sys.exit(1)

        cap = cv2.VideoCapture(args.video_path)
        if not cap.isOpened():
            result['errors'].append(f"Nao consegui abrir video: {args.video_path}")
            print(json.dumps(result))
            sys.exit(1)

        fps = cap.get(cv2.CAP_PROP_FPS)
        frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        duration = frame_count / fps if fps > 0 else 0

        result['metadata'] = {
            'fps': fps,
            'frame_count': frame_count,
            'width': width,
            'height': height,
            'duration': duration,
        }

        opts = {
            'algorithm': args.algorithm,
            'bbox_width': args.bbox_width,
            'bbox_height': args.bbox_height,
            'confidence_threshold': args.confidence_threshold,
            'max_lost_frames': args.max_lost_frames,
            'search_window_scale': args.search_window_scale,
            'window_size': args.window_size,
            'pyramid_levels': args.pyramid_levels,
            'max_iterations': args.max_iterations,
            'tracking_quality': args.tracking_quality,
        }

        for pi, point in enumerate(points):
            cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            if args.algorithm == 'LUCAS_KANADE':
                history, err = track_lucas_kanade(cap, fps, width, height, point, opts)
            else:
                history, err = track_correlation_filter(cap, fps, width, height, point, opts)
            if err:
                result['errors'].append(f"Point {pi}: {err}")
                continue

            # Smoothing
            if args.smoothing == 'whittaker':
                history = apply_whittaker(history, args.whittaker_lambda)
            elif args.smoothing == 'kalman':
                history = apply_kalman(history, args.kalman_process_noise, args.kalman_measurement_noise)
            elif args.smoothing == 'supersmoother':
                history = apply_supersmoother(history, args.supersmoother_alpha)
            # 'none' = skip

            result['tracks'].append({'id': pi, 'history': history})

        cap.release()
        result['success'] = len(result['tracks']) > 0

    except Exception as e:
        result['errors'].append(f"Erro: {str(e)}")
        result['errors'].append(traceback.format_exc())

    print(json.dumps(result))
    sys.exit(0 if result['success'] else 1)


if __name__ == '__main__':
    main()
