# High-Precision 3D Motion Control

A Windows desktop controller for high-precision three-axis motion platforms used in large acoustic-field scanning. It combines TCP motion control, live XYZ position feedback, software travel limits, configurable serpentine scanning, and synchronized RIGOL oscilloscope acquisition.

## Features

- Relative, absolute, and press-and-hold XYZ motion
- Live position feedback and configurable software limits
- Serpentine scans on XY, XZ, or YZ planes
- Independent positive/negative direction selection for both scan axes
- Scan ETA and coordinates relative to the scan start point
- RIGOL waveform acquisition through VISA
- Fast continuous NPY datasets with optional per-point MATLAB-compatible XLSX files
- Scan metadata, validity masks, and actual XYZ positions

## Requirements

- Windows with Python 3.9 or newer
- The motion-controller vendor TCP service
- A VISA runtime such as NI-VISA for oscilloscope access

Install the Python packages:

```powershell
py -m pip install -r requirements.txt
```

## Run

```powershell
py motion_platform_ui_v13_optimized.py
```

Enter the motion service host and port, then replace the oscilloscope VISA placeholder with the resource string reported by your VISA software.

## Safety

Set the hardware limits, zero point, and safe travel area in the vendor software before moving the platform. Verify the software limits against the physical machine; the included values are installation-specific examples, not universal limits.

Run the built-in offline checks with:

```powershell
py motion_platform_ui_v13_optimized.py --self-test
```

## Data output

Each scan can store a compact NPY dataset containing waveforms, timestamps, actual XYZ positions, coordinate axes, and a validity mask. Per-point XLSX output is available for compatibility but is slower.
