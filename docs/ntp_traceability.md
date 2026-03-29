# Analysis of NMI's "Making PC Time Traceable" Procedure

Original document docs\nmi-using-ntp-for-traceable-time-and-frequency.docx

## The Original Text

> *"For a particular day, examine the loopstats file and look at the offsets which are applied – the mean of these could be used as an estimate of the drift of the PC time between synchronizations. Then examine the delays in the peerstats file for those clocks which have been selected as suitable references by ntpd (as indicated by the clock status flags) and take the mean of these as an estimate of the synchronization accuracy. The total uncertainty might then be estimated as the quadrature sum of the mean drift and delay."*

---

## What's Ambiguous

### 1. "Mean of offsets" = "drift between synchronizations"?

The loopstats offset is the **correction applied to the PC clock** at each synchronisation. Calling its mean "drift" conflates two different things:

| Concept | What It Actually Is | Where It Lives |
|---|---|---|
| **Drift** | Rate of clock frequency error (seconds/second) | loopstats **frequency** column |
| **Offset** | How far the clock was from the server at correction time | loopstats **offset** column |
| **Residual wander** | How far the clock drifts between corrections | Derived from offset pattern |

The mean offset tells you the **average residual error** after corrections, not the drift rate. A systematic non-zero mean could indicate a bias, while the spread indicates how much the clock wanders between syncs.

**Also unclear:** Mean of signed offsets or absolute offsets? Signed offsets could cancel out and give a misleadingly small value.

### 2. "Mean of delays" = "synchronization accuracy"?

The peerstats delay is the **round-trip time**. From Application 2 in the same document, the uncertainty from network delay is **half the round-trip delay** (assuming asymmetric worst case), reduced by **1/√3** for a rectangular distribution.

But the procedure says to just take the "mean of delays" — it doesn't mention:
- Halving the delay
- Applying a distribution correction factor
- Whether to use mean, max, or some percentile

### 3. "Quadrature sum of mean drift and delay"

Quadrature sum = √(a² + b²). But:
- Are these **means** (estimates of systematic error)?
- Or **standard deviations** (estimates of random uncertainty)?
- Or **half-ranges** (worst-case bounds)?
- Should the delay be halved first?
- What confidence level does this represent?

---

## NTP Log File Formats

For context, here's what's actually in the files:

### loopstats
```
# MJD    time(s)   offset(s)    freq(ppm)   jitter(s)   allan_dev   time_const
57098    43200     -0.000234    -1.234       0.000156    0.0234      6
```

### peerstats
```
# MJD    time(s)   peer_addr       status  offset(s)   delay(s)    dispersion(s)  jitter(s)
57098    43200     203.0.178.191   9614    -0.000123    0.05692     0.00341        0.000234
```

---

## Possible Mathematical Interpretations

### Interpretation A: Literal Reading (What They Probably Meant)

The simplest reading — means as point estimates, combined in quadrature:

```python
import pandas as pd
import numpy as np

# Load loopstats
loopstats = pd.read_csv('loopstats', sep='\s+',
    names=['mjd', 'time', 'offset', 'freq', 'jitter', 'allan', 'tc'])

# Load peerstats (filtered to selected peers via status flags)
peerstats = pd.read_csv('peerstats', sep='\s+',
    names=['mjd', 'time', 'peer', 'status', 'offset', 'delay', 'dispersion', 'jitter'])

# Filter to a specific day
day = loopstats['mjd'].max()
loop_day = loopstats[loopstats['mjd'] == day]
peer_day = peerstats[peerstats['mjd'] == day]

# Filter peerstats to selected reference clocks
# Status flags: bit 0x0100 indicates sys.peer, 0x0900 = selected
peer_day = peer_day[peer_day['status'].apply(lambda x: int(str(x), 16) & 0x0700 != 0)]

# "Mean of offsets" — literal
mean_offset = loop_day['offset'].mean()

# "Mean of delays"  — literal
mean_delay = peer_day['delay'].mean()

# "Quadrature sum"
uncertainty = np.sqrt(mean_offset**2 + mean_delay**2)

print(f"Mean offset: {mean_offset*1000:.3f} ms")
print(f"Mean delay:  {mean_delay*1000:.3f} ms")
print(f"Uncertainty: {uncertainty*1000:.3f} ms")
```

**Problem:** Signed offsets cancel out. Mean delay isn't halved. No distribution correction. Mixes systematic and random contributions.

---

### Interpretation B: Metrologically Consistent (What They Should Have Meant)

Following the principles from their own Application 2 example:

```python
import pandas as pd
import numpy as np

# --- Component 1: Clock wander between syncs ---
# Use absolute offsets (avoid sign cancellation)
# Mean absolute offset = typical error between corrections
mean_abs_offset = loop_day['offset'].abs().mean()

# Assume rectangular distribution: divide by sqrt(3)
u_offset = mean_abs_offset / np.sqrt(3)

# --- Component 2: Network path uncertainty ---
# Half the round-trip delay is the bound on one-way asymmetry
# (consistent with Application 2 in the same document)
mean_half_delay = peer_day['delay'].mean() / 2

# Assume rectangular distribution: divide by sqrt(3)
u_delay = mean_half_delay / np.sqrt(3)

# --- Component 3: Server uncertainty (from document) ---
u_server = 3e-6  # ±3 µs as stated in document

# --- Combined standard uncertainty ---
u_combined = np.sqrt(u_offset**2 + u_delay**2 + u_server**2)

# --- Expanded uncertainty (k=2, ~95% confidence) ---
U_expanded = 2 * u_combined

print(f"u_offset:  {u_offset*1000:.3f} ms")
print(f"u_delay:   {u_delay*1000:.3f} ms")
print(f"u_server:  {u_server*1e6:.1f} µs")
print(f"U_combined (k=2): ±{U_expanded*1000:.3f} ms")
```

---

### Interpretation C: Statistical (Most Rigorous)

Uses the full distribution of observations rather than just means:

```python
import pandas as pd
import numpy as np

# --- Component 1: Systematic offset (bias) ---
# Mean signed offset = systematic bias in PC time
bias = loop_day['offset'].mean()

# --- Component 2: Random offset variation ---
# Std of offsets = how much the clock wanders between syncs
u_wander = loop_day['offset'].std()

# --- Component 3: Network asymmetry uncertainty ---
# Use mean delay / 2 as the bound
# But also account for variation in delay
mean_half_delay = peer_day['delay'].mean() / 2
std_half_delay = peer_day['delay'].std() / 2

# Rectangular distribution for the asymmetry bound
u_asymmetry = mean_half_delay / np.sqrt(3)

# Add delay variation as additional random component
u_delay_variation = std_half_delay

# --- Component 4: Server uncertainty ---
u_server = 3e-6 / np.sqrt(3)  # ±3 µs rectangular

# --- Combined standard uncertainty ---
u_combined = np.sqrt(u_wander**2 + u_asymmetry**2 + u_delay_variation**2 + u_server**2)

# --- The time on the PC ---
# Corrected time = PC time + bias ± U
# (bias is a known systematic offset that could be corrected for)

U_expanded = 2 * u_combined

print(f"Systematic bias:       {bias*1000:+.3f} ms")
print(f"u_wander:              {u_wander*1000:.3f} ms")
print(f"u_asymmetry:           {u_asymmetry*1000:.3f} ms")
print(f"u_delay_variation:     {u_delay_variation*1000:.3f} ms")
print(f"u_server:              {u_server*1e6:.3f} µs")
print(f"U_expanded (k=2):      ±{U_expanded*1000:.3f} ms")
```

---

### Interpretation D: Use ntpd's Own Statistics

The log files already contain jitter and dispersion — arguably better estimates:

```python
import pandas as pd
import numpy as np

# loopstats already reports RMS jitter = residual sync error
u_jitter = loop_day['jitter'].mean()

# peerstats dispersion = NTP's own estimate of peer uncertainty
u_dispersion = peer_day['dispersion'].mean()

# peerstats jitter = variation in peer offset measurements
u_peer_jitter = peer_day['jitter'].mean()

# Network asymmetry (still need half-delay)
u_asymmetry = peer_day['delay'].mean() / 2 / np.sqrt(3)

# Server uncertainty
u_server = 3e-6 / np.sqrt(3)

# Combined
u_combined = np.sqrt(u_jitter**2 + u_dispersion**2 + u_asymmetry**2 + u_server**2)
U_expanded = 2 * u_combined

print(f"u_jitter (loopstats):  {u_jitter*1e6:.1f} µs")
print(f"u_dispersion:          {u_dispersion*1000:.3f} ms")
print(f"u_peer_jitter:         {u_peer_jitter*1e6:.1f} µs")
print(f"u_asymmetry:           {u_asymmetry*1000:.3f} ms")
print(f"U_expanded (k=2):      ±{U_expanded*1000:.3f} ms")
```

---

## Comparison of Interpretations

| | A (Literal) | B (Metrological) | C (Statistical) | D (NTP native) |
|---|---|---|---|---|
| **Offset handling** | Signed mean | Absolute mean / √3 | Mean (bias) + std (random) | ntpd jitter |
| **Delay handling** | Full mean delay | Half delay / √3 | Half delay / √3 + std | Half delay / √3 + dispersion |
| **Server uncertainty** | Ignored | Included | Included | Included |
| **Distribution assumption** | None | Rectangular | Gaussian + rectangular | NTP's own model |
| **Confidence level** | Undefined | ~68% (k=1) or ~95% (k=2) | ~95% (k=2) | ~95% (k=2) |
| **Sign cancellation risk** | Yes ❌ | No | Separated | No |

---

## Recommendation

> **Use Interpretation B or C** depending on your needs. The original text (Interpretation A) is inconsistent with the document's own Application 2 methodology, which correctly halves the delay and applies a rectangular distribution factor. Interpretation D is attractive but mixes NTP's internal statistics model with metrological uncertainty, which may not satisfy auditors who want explicit uncertainty budgets.

---

## Offset Accuracy to UTC — Practical Variants (E, F, G)

Interpretations A–D estimate the **accuracy of the PC clock** over a whole day and treat the NTP offset as a proxy for clock error. The practical variants below instead focus on a narrower question: **how accurately does the NTP-measured offset represent the true offset to UTC?**

This is the right question when you want to correct a recorded PC timestamp to UTC and state an accuracy, e.g. for an astronomy observation log.

The key difference: these variants do not include clock wander over the day (that is the PC clock accuracy problem). They capture only the uncertainty in the *link* between the NTP server and the PC at the moment of measurement.

### Variant E — Minimal (Network + Measurement)

The two irreducible uncertainty sources when you have NTP running:

| Component | Formula | Notes |
|---|---|---|
| Network asymmetry | $u_{\text{asymmetry}} = \dfrac{\bar{d}/2}{\sqrt{3}}$ | Rectangular distribution on the one-way path split |
| Measurement scatter | $u_{\text{meas}} = \sigma(\text{offset})$ | Standard deviation of loopstats offsets over the dataset |

$$u_{\text{combined}} = \sqrt{u_{\text{asymmetry}}^2 + u_{\text{meas}}^2}$$

$$U_{\text{expanded}} = 2 \, u_{\text{combined}} \quad (k=2,\ \approx\!95\%\ \text{confidence})$$

**When to use:** Best when delays are stable, NTP is disciplined, and you do not need to account for server chain uncertainty.

### Variant F — NTP Dispersion Directly

NTP's own `dispersion` field in peerstats is its internal estimate of total peer uncertainty, incorporating server root dispersion, path delay variations, and timing jitter accumulated along the server chain.

$$u_{\text{offset}} = \overline{\text{dispersion}} \qquad U_{\text{expanded}} = 2 \, u_{\text{offset}}$$

**When to use:** Quick conservative estimate. NTP's dispersion can be larger than necessary (it is a worst-case bound), so this tends to over-estimate. Useful when you want to cite a figure that NTP itself guarantees rather than one derived externally.

**Caution:** Dispersion does not account for systematic network asymmetry directly — it captures accumulated wander along the server chain. For high-asymmetry paths, Variant E may be more representative.

### Variant G — Conservative (Worst-Case Delay)

Same as Variant E but uses the **maximum** observed round-trip delay instead of the mean. This guards against path bursts or transient congestion that the mean conceals.

| Component | Formula |
|---|---|
| Network asymmetry (worst-case) | $u_{\text{asymmetry}} = \dfrac{d_{\max}/2}{\sqrt{3}}$ |
| Measurement scatter | $u_{\text{meas}} = \sigma(\text{offset})$ |

$$U_{\text{expanded}} = 2\sqrt{u_{\text{asymmetry}}^2 + u_{\text{meas}}^2}$$

**When to use:** When reporting accuracy for regulatory or audit purposes where a worst-case bound is required, or when the delay distribution is known to be skewed.

### Variant Comparison

| | E (Minimal) | F (Dispersion) | G (Conservative) |
|---|---|---|---|
| **Delay basis** | Mean RTT | — | Max RTT |
| **Network asymmetry** | ✓ (mean) | Implicit in dispersion | ✓ (max) |
| **Server chain uncertainty** | ✗ | ✓ (via dispersion) | ✗ |
| **Typical result** | Smallest | Mid-range | Largest |
| **Best use case** | Stable, known network | Audit-friendly single figure | Worst-case bound |

---

## Point-in-Time Offset Accuracy Estimate

### Motivation

The interpretations above all compute statistics over a **full day's worth of logs**. For observations (e.g. an astronomical occultation), you need to know: *"what was the NTP offset at 23:14:07 UTC, and how accurate was it?"*

A day-average answer is not satisfying because:
- The offset may have drifted since the last NTP correction
- The relevant peer measurements are those near the event, not from hours earlier

### Method

Given a query time **T** (specified as MJD + seconds-past-midnight):

#### Step 1 — Best-estimate offset at T

Find the loopstats records immediately **before** and **after** T:
- $\hat{\delta}_{\text{before}}$, $t_{\text{before}}$ — offset and timestamp of the last record at or before T
- $\hat{\delta}_{\text{after}}$, $t_{\text{after}}$ — offset and timestamp of the first record after T (if available)
- $f$ — the `freq` correction (ppm) from the record before T

**Case 1 — Both records available (interpolation):**

$$\alpha = \frac{T - t_{\text{before}}}{t_{\text{after}} - t_{\text{before}}}$$

$$\hat{\delta}(T) = \hat{\delta}_{\text{before}} + \alpha\,(\hat{\delta}_{\text{after}} - \hat{\delta}_{\text{before}})$$

Linear interpolation between the two measured offsets is preferred because it uses actual measurements on both sides of T. It avoids reliance on the `freq` field as an offset predictor.

**Why not use `freq` for the projection?** The `freq` column is the correction NTP is *already injecting* into the kernel clock — not an uncompensated drift. Applying it forward as a drift predictor double-counts an effect that is already reflected in the subsequent `offset` reading, and produces spuriously large or small estimates.

**Case 2 — Only a record before T (extrapolation):**

$$\hat{\delta}(T) = \hat{\delta}_{\text{before}}$$

The last known offset is used directly. The `freq` value informs the *uncertainty* on this estimate (see Step 2) but does not shift the best estimate itself.

#### Step 2 — Drift uncertainty

**Case 1 — Interpolating:** The jitter values at the surrounding records bound how much the true offset deviates from a straight line between them:

$$u_{\text{drift}} = \max\!\left(\text{jitter}_{\text{before}},\ \text{jitter}_{\text{after}}\right)$$

**Case 2 — Extrapolating:** The residual frequency error accumulates as uncompensated drift. Modelling this as a rectangular distribution:

$$u_{\text{drift}} = \max\!\left(\frac{|f \times 10^{-6} \times \Delta t|}{\sqrt{3}},\ \text{jitter}_{\text{before}}\right)$$

where $\Delta t = T - t_{\text{before}}$. The loopstats `jitter` is used as a floor in both cases — it represents the irreducible residual error from NTP's own synchronisation noise.

**Note on gap size:** For typical NTP poll intervals (64–1024 s), interpolation is the common case and $u_{\text{drift}}$ is small. For large extrapolation gaps (e.g. NTP was suspended or stepped), the drift term can grow significantly and will dominate the uncertainty budget.

#### Step 3 — Network asymmetry uncertainty

Collect peerstats records for the **active selected server** within a ±1 hour window around T. Compute the mean round-trip delay $\bar{d}$ of those records.

The one-way delay is unknown; the worst-case split is $\bar{d}/2$ in either direction. Assuming a rectangular distribution:

$$u_{\text{asymmetry}} = \frac{\bar{d}/2}{\sqrt{3}}$$

If no records from the active server are found near T, all selected-peer records in the window are used as a fallback.

#### Step 3a — Alternative estimate from candidate peers

NTP disciplines the PC clock to the **sys.peer** (select code 6 or 7) only, so the estimate above is limited by that peer's RTT. Any other peer that has passed NTP's screening — *candidate* (code 4) or *backup* (code 5) — carries its own independently-measured UTC offset in `peerstats` column `[4]`, together with its own RTT and jitter.

For each such peer near T the combined standard uncertainty of using it as a standalone estimate is:

$$u_{\text{cand}} = \sqrt{\left(\frac{d_{\text{cand}}/2}{\sqrt{3}}\right)^2 + \text{jitter}_{\text{cand}}^2}$$

The peer with the lowest $u_{\text{cand}}$ is selected as the **alternative estimate**. If $U_{\text{expanded,cand}} < U_{\text{expanded,sys.peer}}$, it is reported as an improvement and can be used in place of the loopstats-based estimate.

**Important constraint:** the candidate peer's offset is used *together with its own RTT* — never mixed with the loopstats `best_offset`. Substituting another peer's smaller RTT into the loopstats-based uncertainty would be metrologically invalid; the loopstats offset reflects the sys.peer's path specifically.

#### Step 4 — Measurement scatter

Collect loopstats records within ±1 hour of T and compute the sample standard deviation of their offsets:

$$u_{\text{scatter}} = \sigma\!\left(\{\delta_i : |t_i - T| \leq 3600\,\text{s}\}\right)$$

This captures how reproducibly NTP measures the offset around this time — including real network path variation, local clock noise, and any systematic drift not captured by the frequency model.

#### Step 5 — Combined uncertainty

The three components are assumed independent. Applying k=2 (Gaussian) to the full combined standard uncertainty overcounts the asymmetry contribution by ~15% relative to its hard physical ceiling (RTT/2): when asymmetry dominates, the rectangular distribution produces a "95%" interval that exceeds a quantity that already contains 100% of the probability mass. Instead, the asymmetry term and the statistical terms are expanded independently with coverage factors matched to their respective distributions and then combined in quadrature.

$$u_{\text{combined}} = \sqrt{u_{\text{drift}}^2 + u_{\text{asymmetry}}^2 + u_{\text{scatter}}^2}$$

where $u_{\text{asymmetry}} = \bar{d}/2/\sqrt{3}$ is the rectangular standard uncertainty (retained for the k=1 display).

Define the **hard physical bound** on one-way asymmetry and the **Gaussian statistical** component:

$$b_{\text{asym}} = \frac{\bar{d}}{2}$$

$$u_{\text{stat}} = \sqrt{u_{\text{drift}}^2 + u_{\text{scatter}}^2}$$

The expanded uncertainty uses the correct 95% factor for each distribution — 0.95 for rectangular (95% of a uniform U[−b, b] lies within ±0.95b), k=2 for the near-Gaussian statistical terms:

$$U_{\text{expanded}} = \sqrt{(0.95\, b_{\text{asym}})^2 + (2\, u_{\text{stat}})^2} \quad (\approx\!95\%\ \text{confidence})$$

This guarantees $U_{\text{expanded}} \leq b_{\text{asym}} = \bar{d}/2$: the reported interval never exceeds the hard physical ceiling. $u_{\text{combined}}$ (k=1) is retained for display but is not used directly in the expanded coverage calculation.

#### Corrected time

$$t_{\text{UTC}} = t_{\text{PC}} - \hat{\delta}(T) \pm U_{\text{expanded}}$$

where $t_{\text{PC}}$ is the raw PC timestamp at the event.

### Uncertainty Budget Example

For a typical internet NTP setup with ~50 ms RTT (interpolating between records 64 s apart):

| Component | Typical value |
|---|---|
| $u_{\text{drift}}$ (max jitter at surrounding records) | ±0.1–0.5 ms |
| $u_{\text{asymmetry}}$ (RTT = 50 ms) | ±14.4 ms |
| $u_{\text{scatter}}$ (stable NTP) | ±0.1–1 ms |
| **U_expanded (~95%)** | **±24 ms** |

For a local stratum-1 server (~1 ms RTT):

| Component | Typical value |
|---|---|
| $u_{\text{drift}}$ (max jitter at surrounding records) | ±5–50 µs |
| $u_{\text{asymmetry}}$ (RTT = 1 ms) | ±0.29 ms |
| $u_{\text{scatter}}$ | ±0.05–0.2 ms |
| **U_expanded (~95%)** | **±0.5 ms** |

For extrapolation (T is after the last loopstats record, gap = 60 s, freq = 5 ppm):

| Component | Typical value |
|---|---|
| $u_{\text{drift}}$ (freq × gap / √3, floored at jitter) | ±0.17 ms |
| $u_{\text{asymmetry}}$ (RTT = 50 ms) | ±14.4 ms |
| $u_{\text{scatter}}$ | ±0.1–1 ms |
| **U_expanded (~95%)** | **±24 ms** |

The network asymmetry term dominates in all cases. To improve accuracy, reduce RTT (use a nearby server) or use a GPS-disciplined local stratum-1 source.

### Limitations

| Limitation | Impact |
|---|---|
| Linear interpolation model | Assumes offset changes linearly between log records. Real NTP corrections can be non-linear if a step occurred between records; the jitter floor partially accounts for this. |
| Extrapolation at end of data | When T is past the last loopstats record, the best estimate is the last known offset and uncertainty grows with the gap via the `freq` term. |
| Unknown asymmetry distribution | The one-way path split is modelled as rectangular U[−RTT/2, RTT/2]. The 95% coverage factor for a rectangular distribution is 0.95 (not k=2), and this is what the expansion uses. If the path is known to be symmetric (e.g. local LAN), $b_{\text{asym}}$ could be reduced accordingly. |
| NTP polling gaps | If NTP poll intervals are long (64–1024 s), the window between records is wider and interpolation is less precise. Watch `gap_before_s` and `gap_after_s` in the report. |
| Server chain uncertainty | The uncertainty of the reference server itself (typically ±3 µs for a tier-1 server) is not included. It is negligible against network uncertainty in almost all practical cases. |
| Log granularity | loopstats is written once per NTP discipline cycle (typically 64–1024 s). Events between log entries are estimated by interpolation, not direct measurement. |
| Candidate-peer estimate gap | The alternative estimate uses the peerstats record nearest in time to T. If that record is from a long poll ago, the offset shown may not reflect conditions exactly at T. The time gap from T is reported (`alt_gap_s`) so this can be assessed. |
| Candidate peers not always available | Some NTP configurations poll only a single server, or all peers may be rejected/falseticker near T. In this case only the loopstats-based estimate is produced. |

### API Reference

```python
result = estimate_offset_at_time(
    query_mjd,        # int: Modified Julian Day of the event
    query_sec,        # float: seconds past midnight UTC
    loop_rows,        # list of LoopRecord from parse_loopstats()
    peer_rows,        # list of PeerRecord from parse_peerstats()
    window_seconds,   # float: half-width of near-T peer/loop window (default 3600 s)
)

# Key output fields:
result["best_offset"]      # float, seconds: estimated NTP offset at T
                            #   (interpolated if records exist on both sides;
                            #    last known offset if only before-T record available)
result["offset_before"]    # float, seconds: loopstats offset at last record before T
result["offset_after"]     # float or None: loopstats offset at first record after T
result["u_combined"]       # float, seconds: combined standard uncertainty (k=1)
result["u_expanded"]       # float, seconds: expanded uncertainty (k=2, ~95%)
result["u_drift"]          # float, seconds: drift/interpolation uncertainty component
result["u_asymmetry"]      # float, seconds: network asymmetry component
result["u_scatter"]        # float, seconds: measurement scatter component
result["gap_before_s"]     # float, seconds: time since last loopstats record
result["gap_after_s"]      # float or None: time until next loopstats record
result["freq_ppm"]         # float: NTP frequency correction at last record before T
result["active_server_at_T"]       # str: server address active at T
result["mean_delay_near_T"]        # float, seconds: mean RTT of peer records near T
result["n_peer_near_T"]            # int: number of peer records used for network estimate
result["n_candidate_peers_near_T"] # int: distinct candidate-or-better peer servers near T
                                    #   (select code >= 4, including sys.peer)

# Alternative candidate-peer estimate (None when no non-sys.peer candidates are near T):
result["alt_best_offset"]   # float or None: offset from the best candidate peer record near T
result["alt_u_asymmetry"]   # float or None: u_asymmetry for that peer (its RTT/2/sqrt(3))
result["alt_u_scatter"]     # float or None: that peer's jitter (used as u_scatter)
result["alt_u_combined"]    # float or None: sqrt(u_asym^2 + u_scatter^2)
result["alt_u_expanded"]    # float or None: 2 * alt_u_combined (k=2)
result["alt_server"]        # str or None: server address of the best candidate peer
result["alt_delay"]         # float or None: RTT of that peer's record
result["alt_gap_s"]         # float or None: time (s) between T and that peer record
```
