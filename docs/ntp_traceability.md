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
