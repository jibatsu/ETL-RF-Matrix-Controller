#!/usr/bin/env python3
"""
ETL Vortex Matrix Controller
A portable GUI application for controlling ETL Vortex matrix routers.
Built with PySide6 for optimal performance.
"""

import sys
import socket
import threading
import json
import os
import re
import time
from dataclasses import dataclass, field, asdict
from typing import Optional, Dict, List, Tuple
from datetime import datetime

# Fix Windows taskbar icon (must be before QApplication is created)
import platform
if platform.system() == 'Windows':
    try:
        import ctypes
        # Set app user model ID so Windows uses our icon in taskbar
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('com.etl.rt.matrix.controller')
    except Exception:
        pass

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QLineEdit, QSpinBox, QDialog,
    QDialogButtonBox, QGroupBox, QFormLayout, QComboBox, QColorDialog,
    QMessageBox, QFileDialog, QMenuBar, QMenu, QStatusBar, QFrame,
    QSizePolicy, QSpacerItem, QToolBar, QTextEdit, QSplitter, QTabWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QRadioButton
)
from PySide6.QtCore import Qt, Signal, QObject, QTimer, QThread
from PySide6.QtGui import QColor, QFont, QAction, QPalette, QTextCursor

# Check for reset flag on startup
if '--reset' in sys.argv:
    import platform
    if platform.system() == "Darwin":
        config_path = os.path.expanduser("~/Library/Application Support/ETL RF Matrix Controller/etl_config.json")
    elif platform.system() == "Windows":
        config_path = os.path.join(os.environ.get("APPDATA", ""), "ETL RF Matrix Controller", "etl_config.json")
    else:
        config_path = os.path.expanduser("~/.config/etl-rf-matrix-controller/etl_config.json")
    
    if os.path.exists(config_path):
        try:
            os.remove(config_path)
        except:
            pass

def parse_range_string(range_str: str) -> List[int]:
    """Parse a range string like '1-16' or '49-64' or '1,3,5-10' into a list of integers."""
    result = []
    range_str = range_str.strip()
    if not range_str:
        return result
    
    parts = range_str.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            try:
                start, end = part.split('-', 1)
                start, end = int(start.strip()), int(end.strip())
                if start <= end:
                    result.extend(range(start, end + 1))
                else:
                    result.extend(range(start, end - 1, -1))
            except ValueError:
                pass
        else:
            try:
                result.append(int(part))
            except ValueError:
                pass
    
    return result


def format_range_string(numbers: List[int]) -> str:
    """Convert a list of integers to a compact range string, preserving order."""
    if not numbers:
        return ""
    
    # Don't sort - preserve the original order
    ranges = []
    i = 0
    while i < len(numbers):
        start = numbers[i]
        end = start
        
        # Look for consecutive numbers
        while i + 1 < len(numbers) and numbers[i + 1] == end + 1:
            i += 1
            end = numbers[i]
        
        if start == end:
            ranges.append(str(start))
        else:
            ranges.append(f"{start}-{end}")
        i += 1
    
    return ", ".join(ranges)


@dataclass
class OutputGroup:
    name: str
    color: str
    outputs: List[int] = field(default_factory=list)


@dataclass
class RoutePreset:
    """A saved routing preset/scene."""
    name: str
    routes: Dict[int, int]  # output -> input mapping
    outputs: Optional[List[int]] = None  # If None, applies to all outputs; otherwise specific outputs only
    router_name: str = ""

    def to_dict(self):
        return {
            'name': self.name,
            'routes': {str(k): v for k, v in self.routes.items()},
            'outputs': self.outputs
        }
    
    @classmethod
    def from_dict(cls, d):
        return cls(
            name=d.get('name', 'Unnamed'),
            routes={int(k): v for k, v in d.get('routes', {}).items()},
            outputs=d.get('outputs')
        )


@dataclass
class RouterConfig:
    ip_address: str = ""
    port: int = 4000
    num_inputs: int = 0
    num_outputs: int = 0
    primary_num_outputs: int = 0  # Actual output count of primary router
    router_name: str = ""
    input_names: Dict[int, str] = field(default_factory=dict)
    output_groups: List[OutputGroup] = field(default_factory=list)
    button_labels: Dict[str, str] = field(default_factory=dict)
    first_run: bool = True
    label_font_family: str = "Helvetica"
    label_font_size: int = 10
    button_font_family: str = "Helvetica"
    button_font_size: int = 9
    active_route_color: str = "#83f600"
    # Show numbers next to labels
    show_input_numbers: bool = False
    show_output_numbers: bool = False
    dark_theme: bool = True  # True = dark theme, False = light theme
    # Crosshair hover effect
    crosshair_enabled: bool = False
    crosshair_luminance_shift: int = 20  # Percentage to shift luminance (0-50)
    crosshair_border_color: str = "#ffffff"  # Border color for crosshair
    # Row luminance adjustments (input_num -> luminance_shift percentage, -50 to +50)
    row_luminance: Dict[int, int] = field(default_factory=dict)
    # Route presets/scenes
    route_presets: List[RoutePreset] = field(default_factory=list)
    # Compact mode
    compact_mode: bool = False
    # Toolbar visibility
    show_toolbar: bool = True
    toolbar_buttons_visible: Dict[str, bool] = field(default_factory=lambda: {
        'settings': True,
        'refresh': False,
        'telemetry': True,
        'presets': True,
        'compact': False,
        'fit': True,
        'connection': True
    })
    # Advanced: custom input/output ranges
    use_custom_ranges: bool = False
    custom_inputs: List[int] = field(default_factory=list)
    custom_outputs: List[int] = field(default_factory=list)
    # Combine multiple routers
    combine_routers: bool = False
    additional_routers: List[dict] = field(default_factory=list)  # [{'ip': '...', 'port': 4001, 'num_outputs': 16}, ...]
    

    def to_dict(self):
        d = asdict(self)
        # Properly serialize route_presets
        d['route_presets'] = [p.to_dict() if isinstance(p, RoutePreset) else p for p in self.route_presets]
        return d
    
    def get_inputs(self) -> List[int]:
        """Get the list of input numbers to display."""
        if self.use_custom_ranges and self.custom_inputs:
            return self.custom_inputs
        return list(range(1, self.num_inputs + 1))
    
    def get_outputs(self) -> List[int]:
        """Get the list of output numbers to display."""
        if self.use_custom_ranges and self.custom_outputs:
            return self.custom_outputs
        
        # When combining routers, num_outputs is the TOTAL across all routers
        # When not combining, num_outputs is just the primary router's outputs
        return list(range(1, self.num_outputs + 1))
    
    def get_router_for_output(self, output: int) -> Tuple[int, int]:
        """
        Given a combined output number, return (router_index, local_output).
        router_index 0 = primary router, 1+ = additional routers.
        """
        if not self.combine_routers or not self.additional_routers:
            return (0, output)
        
        # Primary router's output count
        primary_outputs = self.primary_num_outputs or 16  # Fallback
        
        # Check if output is on primary router
        if output <= primary_outputs:
            return (0, output)
        
        # Check additional routers
        running_total = primary_outputs
        for i, router in enumerate(self.additional_routers):
            router_outputs = router.get('num_outputs', 0)
            if router_outputs == 0:
                continue
            if output <= running_total + router_outputs:
                local_output = output - running_total
                return (i + 1, local_output)
            running_total += router_outputs
        
        # Output is beyond all defined routers - this means user requested
        # more outputs than routers can provide. Route to last router.
        if self.additional_routers:
            last_router_idx = len(self.additional_routers)
            last_offset = primary_outputs
            for i, router in enumerate(self.additional_routers[:-1]):
                last_offset += router.get('num_outputs', 0)
            local_output = output - last_offset
            return (last_router_idx, local_output)
        
        return (0, output)
    
    def get_display_groups(self) -> List['OutputGroup']:
        """Get groups filtered for current display outputs, preserving custom order."""
        outputs_to_show = self.get_outputs()  # Keep as list to preserve order
        outputs_set = set(outputs_to_show)
        
        # Track which outputs are covered by existing groups
        covered_outputs = set()
        group_map = {}  # output -> group
        
        for group in self.output_groups:
            for out in group.outputs:
                if out in outputs_set:
                    group_map[out] = group
                    covered_outputs.add(out)
        
        # Build display groups in the order of outputs_to_show
        display_groups = []
        current_group = None
        current_outputs = []
        
        for out in outputs_to_show:
            if out in group_map:
                group = group_map[out]
                # Check if this continues the current group
                if current_group is not None and current_group.name == group.name and current_group.color == group.color:
                    current_outputs.append(out)
                else:
                    # Save previous group if any
                    if current_outputs:
                        display_groups.append(OutputGroup(current_group.name, current_group.color, current_outputs))
                    current_group = group
                    current_outputs = [out]
            else:
                # Uncovered output - save current group and add individual
                if current_outputs:
                    display_groups.append(OutputGroup(current_group.name, current_group.color, current_outputs))
                    current_group = None
                    current_outputs = []
                display_groups.append(OutputGroup(f"Out {out}", "#b0b0b0", [out]))
        
        # Don't forget the last group
        if current_outputs:
            display_groups.append(OutputGroup(current_group.name, current_group.color, current_outputs))
        
        return display_groups

    @classmethod
    def from_dict(cls, d):
        config = cls(
            ip_address=d.get('ip_address', ''),
            port=d.get('port', 4000),
            num_inputs=d.get('num_inputs', 0),
            num_outputs=d.get('num_outputs', 0),
            primary_num_outputs=d.get('primary_num_outputs', 0),
            router_name=d.get('router_name', ''),
            input_names={int(k): v for k, v in d.get('input_names', {}).items()},
            button_labels={str(k): v for k, v in d.get('button_labels', {}).items()},
            first_run=d.get('first_run', True),
            label_font_family=d.get('label_font_family', 'Helvetica'),
            label_font_size=d.get('label_font_size', 10),
            button_font_family=d.get('button_font_family', 'Helvetica'),
            button_font_size=d.get('button_font_size', 9),
            active_route_color=d.get('active_route_color', "#83f600"),
            show_input_numbers=d.get('show_input_numbers', False),
            show_output_numbers=d.get('show_output_numbers', False),
            dark_theme=d.get('dark_theme', True),
            crosshair_enabled=d.get('crosshair_enabled', False),
            crosshair_luminance_shift=d.get('crosshair_luminance_shift', 20),
            crosshair_border_color=d.get('crosshair_border_color', '#ffffff'),
            row_luminance={int(k): v for k, v in d.get('row_luminance', {}).items()},
            compact_mode=d.get('compact_mode', False),
            show_toolbar=d.get('show_toolbar', True),
            toolbar_buttons_visible=d.get('toolbar_buttons_visible', {
                'settings': True,
                'refresh': False,
                'telemetry': True,
                'presets': True,
                'compact': False,
                'fit': True,
                'connection': True
            }),
            use_custom_ranges=d.get('use_custom_ranges', False),
            custom_inputs=d.get('custom_inputs', []),
            custom_outputs=d.get('custom_outputs', []),
            combine_routers=d.get('combine_routers', False),
            additional_routers=d.get('additional_routers', []),
        )
        for g in d.get('output_groups', []):
            config.output_groups.append(OutputGroup(g['name'], g['color'], g['outputs']))
        for p in d.get('route_presets', []):
            config.route_presets.append(RoutePreset.from_dict(p))
        return config


class ETLProtocol:
    def __init__(self, ip: str, port: int = 4000, timeout: float = 5.0):
        self.ip = ip
        self.port = port
        self.timeout = timeout
        self._lock = threading.Lock()  # Serialize all router communications

    def _calculate_checksum(self, command: str) -> str:
        """
        Calculate checksum for ETL protocol.
        
        The checksum is XOR of all bytes in the message, then XOR with a 
        command-type-specific key. Keys were derived from packet capture analysis.
        """
        xor_all = 0
        for c in command:
            xor_all ^= ord(c)
        
        content = command[1:-1] if command.startswith('{') and command.endswith('}') else command
        
        # Apply command-specific XOR key based on command type
        if content.startswith('ABc') and ',' in content:
            parts = content.split(',')
            if len(parts) >= 4:  # ABcX,00,00,01 or ABcX,00,00,02 format
                xor_all ^= 0x33  # Key for telemetry with 4 params
            else:  # ABcC,00,00 format (3 params)
                xor_all ^= 0x78  # Key for chassis telemetry
        elif content.startswith('*'):
            xor_all ^= 0x48  # Key for device info (*BI)
        elif content.startswith('ABM'):
            xor_all ^= 0x3D  # Key for matrix config
        elif content.startswith('ABJ'):
            xor_all ^= 0x47  # Key for ABJ commands
        elif content == 'AB?':
            xor_all ^= 0x46  # Key for status query
        elif content.startswith('ABs,'):
            xor_all ^= 0x06  # Key for routing commands
        
        return chr(xor_all & 0x7F)

    def _send_command(self, command: str) -> Optional[str]:
        """Send a command and wait for response with proper timing."""
        with self._lock:  # Serialize router communications
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(self.timeout)
                sock.connect((self.ip, self.port))
                
                # Send the command with checksum
                full_command = command + self._calculate_checksum(command)
                sock.sendall(full_command.encode('ascii'))
                
                # Wait for response - the router may need time to process
                # Read response in a loop until we get a complete message or timeout
                response = b''
                start_time = time.time()
                
                while time.time() - start_time < self.timeout:
                    try:
                        chunk = sock.recv(4096)
                        if chunk:
                            response += chunk
                            # Check if we have a complete response (ends with })
                            decoded = response.decode('ascii', errors='replace')
                            if '}' in decoded:
                                # Got complete response
                                break
                        else:
                            # Connection closed by server
                            break
                    except socket.timeout:
                        break
                
                sock.close()
                
                if response:
                    return response.decode('ascii', errors='replace')
                return None
                
            except Exception as e:
                print(f"Connection error: {e}")
                return None

    def get_device_info(self) -> Optional[str]:
        response = self._send_command("{*BI}")
        if response:
            match = re.search(r'\{BBI,([^,]+),([^}]+)\}', response)
            if match:
                return f"{match.group(1)} - {match.group(2)}"
        return None

    def get_matrix_config(self) -> Optional[Tuple[int, int]]:
        response = self._send_command("{ABM?}")
        if response:
            match = re.search(r'\{BAM\?,(\d+),(\d+)', response)
            if match:
                return (int(match.group(1)), int(match.group(2)))
        return None

    def _calculate_route_checksum(self, output_num: int, input_num: int) -> str:
        """Calculate checksum specifically for routing commands.
        
        The checksum is based on the sum of individual digits in the
        3-digit output and input numbers, plus 106, with wrapping.
        When value > 126, wrap to ASCII 32+ (space and punctuation).
        """
        # Format as 3-digit strings: output=1 → "001", input=5 → "005"
        out_str = f"{output_num:03d}"
        inp_str = f"{input_num:03d}"
        
        # Sum the individual digit values: 0+0+1+0+0+5 = 6
        digit_sum = sum(int(d) for d in out_str + inp_str)
        
        # Add base value of 106 and calculate checksum
        val = 106 + digit_sum
        
        # Wrap to stay in printable ASCII range
        # When > 126, wrap to 32+ (space, !, ", etc.)
        if val > 126:
            val = val - 95  # 127 -> 32, 128 -> 33, etc.
        
        return chr(val)

    def route(self, input_num: int, output_num: int) -> bool:
        """Route an input to an output. 
        
        The router command format is {ABs,OUTPUT,INPUT} - output comes first!
        
        Note: The router executes routing commands but may not send a response.
        We send the command, wait briefly, then return True if no error occurred.
        The actual routing can be verified via get_status().
        """
        with self._lock:  # Serialize router communications
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(2.0)  # Shorter timeout for routing
                sock.connect((self.ip, self.port))
                
                # IMPORTANT: Router format is {ABs,OUTPUT,INPUT} - output first!
                command = f"{{ABs,{output_num:03d},{input_num:03d}}}"
                checksum = self._calculate_route_checksum(output_num, input_num)
                full_command = command + checksum
                sock.sendall(full_command.encode('ascii'))
                
                # Wait briefly for router to process (it may or may not respond)
                response = b''
                try:
                    sock.settimeout(1.0)  # Short timeout - router often doesn't respond
                    while True:
                        chunk = sock.recv(1024)
                        if not chunk:
                            break
                        response += chunk
                        if b'}' in response:
                            break
                except socket.timeout:
                    pass  # Expected - router often doesn't respond to routing
                
                sock.close()
                
                # If we got a response, check it
                if response:
                    response_str = response.decode('ascii', errors='ignore')
                    if 'BAs?' in response_str:
                        return True
                
                # No response is also OK - router executes but doesn't always respond
                # The routing likely worked; caller can verify via get_status()
                return True
                
            except Exception as e:
                print(f"Route error: {e}")
                return False

    def get_status(self) -> Optional[str]:
        return self._send_command("{AB?}")
    
    def get_matrix_telemetry(self, card: int = 0, slot: int = 0) -> Optional[str]:
        return self._send_command(f"{{ABcM,{card:02d},{slot:02d},01}}")
    
    def get_output_telemetry(self, card: int = 0, slot: int = 0) -> Optional[str]:
        return self._send_command(f"{{ABcO,{card:02d},{slot:02d},01}}")
    
    def get_input_telemetry(self, card: int = 0, slot: int = 0) -> Optional[str]:
        return self._send_command(f"{{ABcI,{card:02d},{slot:02d},02}}")
    
    def get_chassis_telemetry(self) -> Optional[str]:
        return self._send_command("{ABcC,00,00}")


class TelemetrySignals(QObject):
    data_received = Signal(str, str)
    status_received = Signal(object)  # Use object instead of dict to avoid PySide6 C++ conversion errors
    error = Signal(str)


class TelemetryThread(QThread):
    def __init__(self, protocol: ETLProtocol, interval: float = 2.0):
        super().__init__()
        self.protocol = protocol
        self.interval = interval
        self.running = False
        self.signals = TelemetrySignals()
        self.poll_status = True
        self.poll_matrix = True
        self.poll_chassis = True
    
    def run(self):
        self.running = True
        while self.running:
            try:
                if self.poll_status:
                    response = self.protocol.get_status()
                    if response:
                        self.signals.data_received.emit("STATUS", response)
                        self._parse_status(response)
                
                if self.poll_matrix:
                    response = self.protocol.get_matrix_telemetry()
                    if response:
                        self.signals.data_received.emit("MATRIX", response)
                
                if self.poll_chassis:
                    response = self.protocol.get_chassis_telemetry()
                    if response:
                        self.signals.data_received.emit("CHASSIS", response)
                
            except Exception as e:
                self.signals.error.emit(str(e))
            
            for _ in range(int(self.interval * 10)):
                if not self.running:
                    break
                time.sleep(0.1)
    
    def _parse_status(self, response: str):
        match = re.search(r'\{BASTATUS,([^}]+)\}', response)
        if match:
            parts = match.group(1).split(',')
            routes = {}
            for i, part in enumerate(parts):
                if part.isdigit():
                    routes[i + 1] = int(part)
            self.signals.status_received.emit(routes)
    
    def stop(self):
        self.running = False


class TelemetryWindow(QDialog):
    def __init__(self, parent, protocol: ETLProtocol):
        super().__init__(parent)
        self.protocol = protocol
        self.telemetry_thread = None
        
        self.setWindowTitle("Telemetry Monitor")
        self.resize(700, 500)
        self.setModal(False)
        
        self._setup_ui()
        self._start_monitoring()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        
        control_bar = QHBoxLayout()
        
        self.status_check = QCheckBox("Routing Status")
        self.status_check.setChecked(True)
        self.status_check.stateChanged.connect(self._update_polling)
        control_bar.addWidget(self.status_check)
        
        self.matrix_check = QCheckBox("Matrix Card")
        self.matrix_check.setChecked(True)
        self.matrix_check.stateChanged.connect(self._update_polling)
        control_bar.addWidget(self.matrix_check)
        
        self.chassis_check = QCheckBox("Chassis")
        self.chassis_check.setChecked(True)
        self.chassis_check.stateChanged.connect(self._update_polling)
        control_bar.addWidget(self.chassis_check)
        
        control_bar.addStretch()
        
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 60)
        self.interval_spin.setValue(2)
        self.interval_spin.setSuffix(" sec")
        self.interval_spin.valueChanged.connect(self._update_interval)
        control_bar.addWidget(QLabel("Interval:"))
        control_bar.addWidget(self.interval_spin)
        
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self._clear_log)
        control_bar.addWidget(self.clear_btn)
        
        layout.addLayout(control_bar)
        
        self.tabs = QTabWidget()
        
        self.raw_log = QTextEdit()
        self.raw_log.setReadOnly(True)
        self.raw_log.setFont(QFont("Courier New", 10))
        self.tabs.addTab(self.raw_log, "Raw Log")
        
        self.status_table = QTableWidget()
        self.status_table.setColumnCount(2)
        self.status_table.setHorizontalHeaderLabels(["Output", "Routed Input"])
        self.status_table.horizontalHeader().setStretchLastSection(True)
        self.tabs.addTab(self.status_table, "Routing Status")
        
        self.chassis_table = QTableWidget()
        self.chassis_table.setColumnCount(2)
        self.chassis_table.setHorizontalHeaderLabels(["Parameter", "Value"])
        self.chassis_table.horizontalHeader().setStretchLastSection(True)
        self.tabs.addTab(self.chassis_table, "Chassis Info")
        
        layout.addWidget(self.tabs)
        
        self.status_label = QLabel("Monitoring...")
        layout.addWidget(self.status_label)
    
    def _start_monitoring(self):
        self.telemetry_thread = TelemetryThread(self.protocol, self.interval_spin.value())
        self.telemetry_thread.signals.data_received.connect(self._on_data_received)
        self.telemetry_thread.signals.status_received.connect(self._on_status_received)
        self.telemetry_thread.signals.error.connect(self._on_error)
        self.telemetry_thread.start()
    
    def _update_polling(self):
        if self.telemetry_thread:
            self.telemetry_thread.poll_status = self.status_check.isChecked()
            self.telemetry_thread.poll_matrix = self.matrix_check.isChecked()
            self.telemetry_thread.poll_chassis = self.chassis_check.isChecked()
    
    def _update_interval(self, value):
        if self.telemetry_thread:
            self.telemetry_thread.interval = value
    
    def _clear_log(self):
        self.raw_log.clear()
    
    def _on_data_received(self, cmd_type: str, data: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.raw_log.append(f"[{timestamp}] {cmd_type}: {data.strip()}")
        
        cursor = self.raw_log.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.raw_log.setTextCursor(cursor)
        
        if cmd_type == "CHASSIS":
            self._parse_chassis(data)
        
        self.status_label.setText(f"Last update: {timestamp}")
    
    def _on_status_received(self, routes: dict):
        self.status_table.setRowCount(len(routes))
        for row, (output, input_num) in enumerate(sorted(routes.items())):
            self.status_table.setItem(row, 0, QTableWidgetItem(f"Output {output}"))
            self.status_table.setItem(row, 1, QTableWidgetItem(f"Input {input_num}"))
    
    def _parse_chassis(self, data: str):
        match = re.search(r'\{BAcC,\d+,\d+,([^}]+)\}', data)
        if match:
            content = match.group(1)
            # Parse the structured data:
            # OSO+320O+300O+291OOO20460O06060O06150O06150O22485O
            # [0:3] = status flags (OSO - O=OK/Open, S=Shut)
            # Temperatures are +XXX format (divide by 10 for °C)
            # 5-digit numbers are fan pulses/min (raw value, no division)
            
            self.chassis_table.setRowCount(0)
            rows = []
            
            # Parse temperatures (+XXX format, divide by 10)
            temp_matches = re.findall(r'[+\-](\d{3})(?=O)', content)
            temp_names = ["CPU Temperature", "PSU 1 Temperature", "PSU 2 Temperature"]
            for i, temp in enumerate(temp_matches[:3]):
                try:
                    temp_c = int(temp) / 10.0
                    name = temp_names[i] if i < len(temp_names) else f"Temperature {i+1}"
                    rows.append((name, f"{temp_c:.1f}°C"))
                except:
                    pass
            
            # Parse fan speeds (5-digit numbers = pulses/min, raw values)
            # Find all 5-digit numbers after the temperature section
            fan_section = re.search(r'OOO(.+)$', content)
            if fan_section:
                fan_data = fan_section.group(1)
                fan_matches = re.findall(r'(\d{5})O', fan_data)
                fan_names = ["Left Fan", "Rear Fan 1", "Rear Fan 2", "Rear Fan 3", "Right Fan"]
                for i, pulses in enumerate(fan_matches[:5]):
                    try:
                        pulses_val = int(pulses)
                        name = fan_names[i] if i < len(fan_names) else f"Fan {i+1}"
                        if pulses_val > 0:
                            rows.append((name, f"{pulses_val} pulses/min"))
                        else:
                            rows.append((name, "Off"))
                    except:
                        pass
            
            # Parse door/status (second character 'S' means Shut, 'O' means Open)
            if len(content) >= 3:
                status_flags = content[0:3]
                # Position [1] appears to be door status: S=Shut, O=Open
                door_char = status_flags[1] if len(status_flags) > 1 else 'O'
                door_status = "Shut" if door_char == 'S' else "Open"
                rows.append(("Rear Door", door_status))
            
            # Populate table
            self.chassis_table.setRowCount(len(rows))
            for row, (param, val) in enumerate(rows):
                self.chassis_table.setItem(row, 0, QTableWidgetItem(param))
                self.chassis_table.setItem(row, 1, QTableWidgetItem(val))
    
    def _on_error(self, error: str):
        self.status_label.setText(f"Error: {error}")
    
    def closeEvent(self, event):
        if self.telemetry_thread:
            self.telemetry_thread.stop()
            self.telemetry_thread.wait(2000)
        event.accept()


class RouteSignals(QObject):
    route_complete = Signal(int, int, bool)


class SettingsDialog(QDialog):
    def __init__(self, parent, config: RouterConfig):
        super().__init__(parent)
        self.config = config
        self.setWindowTitle("Settings")
        self.setMinimumWidth(500)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Create tab widget
        tabs = QTabWidget()
        layout.addWidget(tabs)
        
        # === ROUTER TAB ===
        router_tab = QWidget()
        router_layout = QVBoxLayout(router_tab)
        
        # Connection group
        conn_group = QGroupBox("Connection")
        conn_layout = QFormLayout()
        
        self.ip_edit = QLineEdit(self.config.ip_address)
        ip_row = QHBoxLayout()
        ip_row.addWidget(self.ip_edit)
        self.test_btn = QPushButton("Test")
        self.test_btn.clicked.connect(self._test_connection)
        ip_row.addWidget(self.test_btn)
        conn_layout.addRow("Router IP:", ip_row)
        
        self.port_spin = QSpinBox()
        self.port_spin.setRange(1, 65535)
        self.port_spin.setValue(self.config.port)
        
        self.conn_status = QLabel("")
        conn_layout.addRow("", self.conn_status)
        conn_group.setLayout(conn_layout)
        router_layout.addWidget(conn_group)

        # Matrix size group
        matrix_group = QGroupBox("Matrix Size")
        matrix_layout = QFormLayout()
        
        self.inputs_spin = QSpinBox()
        self.inputs_spin.setRange(1, 256)
        self.inputs_spin.setValue(self.config.num_inputs or 8)
        matrix_layout.addRow("Inputs:", self.inputs_spin)
        
        self.outputs_spin = QSpinBox()
        self.outputs_spin.setRange(1, 256)
        self.outputs_spin.setValue(self.config.num_outputs or 8)
        matrix_layout.addRow("Outputs:", self.outputs_spin)
        
        self.detect_btn = QPushButton("Auto-Detect")
        self.detect_btn.clicked.connect(self._auto_detect)
        matrix_layout.addRow("", self.detect_btn)
        matrix_group.setLayout(matrix_layout)
        router_layout.addWidget(matrix_group)

        # Advanced: Custom ranges
        advanced_group = QGroupBox("Custom Input/Output Ranges")
        advanced_layout = QVBoxLayout()
        
        self.use_custom_check = QCheckBox("Use custom ranges instead of sequential numbering")
        self.use_custom_check.setChecked(self.config.use_custom_ranges)
        self.use_custom_check.stateChanged.connect(self._toggle_custom_ranges)
        advanced_layout.addWidget(self.use_custom_check)
        
        hint_label = QLabel("Specify ranges like: 1-16 or 49-64 or 1,3,5-10,20. Ranges can also be reordered e.g. 1-5,11-20,6-10")
        hint_label.setStyleSheet("color: gray; font-size: 10px;")
        hint_label.setWordWrap(True)
        advanced_layout.addWidget(hint_label)
        
        ranges_form = QFormLayout()
        
        self.custom_inputs_edit = QLineEdit()
        if self.config.custom_inputs:
            self.custom_inputs_edit.setText(format_range_string(self.config.custom_inputs))
        else:
            self.custom_inputs_edit.setPlaceholderText("e.g., 1-48")
        ranges_form.addRow("Input range:", self.custom_inputs_edit)
        
        self.custom_outputs_edit = QLineEdit()
        if self.config.custom_outputs:
            self.custom_outputs_edit.setText(format_range_string(self.config.custom_outputs))
        else:
            self.custom_outputs_edit.setPlaceholderText("e.g., 49-64")
        ranges_form.addRow("Output range:", self.custom_outputs_edit)
        
        advanced_layout.addLayout(ranges_form)
        
        self.range_preview = QLabel("")
        self.range_preview.setStyleSheet("color: #666;")
        advanced_layout.addWidget(self.range_preview)
        
        self.custom_inputs_edit.textChanged.connect(self._update_range_preview)
        self.custom_outputs_edit.textChanged.connect(self._update_range_preview)
        
        advanced_group.setLayout(advanced_layout)
        router_layout.addWidget(advanced_group)
        
        router_layout.addStretch()
        tabs.addTab(router_tab, "Router")
        
        # === APPEARANCE TAB ===
        appearance_tab = QWidget()
        appearance_layout = QVBoxLayout(appearance_tab)
        
        # Theme selection
        theme_group = QGroupBox("Theme")
        theme_layout = QHBoxLayout()
        theme_layout.addWidget(QLabel("Theme:"))
        
        self.dark_theme_radio = QRadioButton("Dark")
        self.light_theme_radio = QRadioButton("Light")
        if self.config.dark_theme:
            self.dark_theme_radio.setChecked(True)
        else:
            self.light_theme_radio.setChecked(True)
        
        theme_layout.addWidget(self.dark_theme_radio)
        theme_layout.addWidget(self.light_theme_radio)
        theme_layout.addStretch()
        theme_group.setLayout(theme_layout)
        appearance_layout.addWidget(theme_group)

        # Active route color
        color_group = QGroupBox("Active Route Highlight")
        color_layout = QHBoxLayout()
        color_layout.addWidget(QLabel("Color:"))
        
        self.color_preview = QLabel("    ")
        self.color_preview.setAutoFillBackground(True)
        self._set_color_preview(self.config.active_route_color)
        color_layout.addWidget(self.color_preview)
        
        self.color_btn = QPushButton("Choose...")
        self.color_btn.clicked.connect(self._choose_color)
        color_layout.addWidget(self.color_btn)
        color_layout.addStretch()
        color_group.setLayout(color_layout)
        appearance_layout.addWidget(color_group)

        # Label numbers
        numbers_group = QGroupBox("Label Numbers")
        numbers_layout = QVBoxLayout()
        
        self.show_input_numbers_check = QCheckBox("Show input numbers column")
        self.show_input_numbers_check.setChecked(self.config.show_input_numbers)
        numbers_layout.addWidget(self.show_input_numbers_check)
        
        self.show_output_numbers_check = QCheckBox("Show output numbers in group headers")
        self.show_output_numbers_check.setChecked(self.config.show_output_numbers)
        numbers_layout.addWidget(self.show_output_numbers_check)
        
        numbers_group.setLayout(numbers_layout)
        appearance_layout.addWidget(numbers_group)

        # Hover/Crosshair settings
        crosshair_group = QGroupBox("Hover Effect")
        crosshair_layout = QVBoxLayout()
        
        crosshair_form = QFormLayout()
        
        self.crosshair_lum_spin = QSpinBox()
        self.crosshair_lum_spin.setRange(-50, 50)
        self.crosshair_lum_spin.setValue(self.config.crosshair_luminance_shift)
        self.crosshair_lum_spin.setSuffix("%")
        crosshair_form.addRow("Hover brightness:", self.crosshair_lum_spin)
        
        crosshair_layout.addLayout(crosshair_form)
        
        self.crosshair_check = QCheckBox("Enable crosshair (highlight full row and column)")
        self.crosshair_check.setChecked(self.config.crosshair_enabled)
        self.crosshair_check.stateChanged.connect(self._toggle_crosshair_settings)
        crosshair_layout.addWidget(self.crosshair_check)
        
        crosshair_border_form = QFormLayout()
        crosshair_border_row = QHBoxLayout()
        self.crosshair_border_preview = QLabel("    ")
        self.crosshair_border_preview.setAutoFillBackground(True)
        self._set_crosshair_border_preview(self.config.crosshair_border_color)
        crosshair_border_row.addWidget(self.crosshair_border_preview)
        
        self.crosshair_border_btn = QPushButton("Choose...")
        self.crosshair_border_btn.clicked.connect(self._choose_crosshair_border_color)
        crosshair_border_row.addWidget(self.crosshair_border_btn)
        crosshair_border_row.addStretch()
        crosshair_border_form.addRow("Crosshair border color:", crosshair_border_row)
        
        crosshair_layout.addLayout(crosshair_border_form)
        crosshair_group.setLayout(crosshair_layout)
        appearance_layout.addWidget(crosshair_group)
        
        self.crosshair_border_color = self.config.crosshair_border_color
        
        appearance_layout.addStretch()
        tabs.addTab(appearance_tab, "Appearance")
        
        # === FONTS TAB ===
        fonts_tab = QWidget()
        fonts_layout = QVBoxLayout(fonts_tab)

        # Matrix Label font group
        label_font_group = QGroupBox("Matrix Labels (Headers and Row Labels)")
        label_font_layout = QFormLayout()
        
        self.label_font_combo = QComboBox()
        self.label_font_combo.addItems(["Helvetica", "Arial", "Verdana", "Tahoma", "Courier New", "Monaco", "Menlo"])
        self.label_font_combo.setCurrentText(self.config.label_font_family)
        self.label_font_combo.setEditable(True)
        label_font_layout.addRow("Family:", self.label_font_combo)
        
        self.label_font_size = QSpinBox()
        self.label_font_size.setRange(6, 24)
        self.label_font_size.setValue(self.config.label_font_size)
        label_font_layout.addRow("Size:", self.label_font_size)
        label_font_group.setLayout(label_font_layout)
        fonts_layout.addWidget(label_font_group)

        # Matrix Button font group
        btn_font_group = QGroupBox("Matrix Buttons")
        btn_font_layout = QFormLayout()
        
        self.btn_font_combo = QComboBox()
        self.btn_font_combo.addItems(["Helvetica", "Arial", "Verdana", "Tahoma", "Courier New", "Monaco", "Menlo"])
        self.btn_font_combo.setCurrentText(self.config.button_font_family)
        self.btn_font_combo.setEditable(True)
        btn_font_layout.addRow("Family:", self.btn_font_combo)
        
        self.btn_font_size = QSpinBox()
        self.btn_font_size.setRange(6, 24)
        self.btn_font_size.setValue(self.config.button_font_size)
        btn_font_layout.addRow("Size:", self.btn_font_size)
        btn_font_group.setLayout(btn_font_layout)
        fonts_layout.addWidget(btn_font_group)
        
        fonts_layout.addStretch()
        tabs.addTab(fonts_tab, "Fonts")

        # === DIALOG BUTTONS ===
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        # Initialize states
        self.active_color = self.config.active_route_color
        self._toggle_custom_ranges()
        self._update_range_preview()
        self._toggle_crosshair_settings()

        # Style group box headers only (using font on title, not contents)
        for group_box in self.findChildren(QGroupBox):
            group_box.setStyleSheet("""
                QGroupBox {
                    font-weight: bold;
                    font-size: 14px;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    padding: 0px 6px;
                }
                QGroupBox QWidget {
                    font-weight: normal;
                    font-size: 12px;
                }
            """)

    def _toggle_custom_ranges(self):
        enabled = self.use_custom_check.isChecked()
        self.custom_inputs_edit.setEnabled(enabled)
        self.custom_outputs_edit.setEnabled(enabled)
        self._update_range_preview()
    
    def _update_range_preview(self):
        if not self.use_custom_check.isChecked():
            self.range_preview.setText("")
            return
        
        inputs = parse_range_string(self.custom_inputs_edit.text())
        outputs = parse_range_string(self.custom_outputs_edit.text())
        
        if inputs and outputs:
            self.range_preview.setText(f"Will show {len(inputs)} inputs × {len(outputs)} outputs")
        elif inputs:
            self.range_preview.setText(f"Will show {len(inputs)} inputs (specify outputs)")
        elif outputs:
            self.range_preview.setText(f"Will show {len(outputs)} outputs (specify inputs)")
        else:
            self.range_preview.setText("Enter ranges to see preview")

    def _set_color_preview(self, color: str):
        palette = self.color_preview.palette()
        palette.setColor(QPalette.Window, QColor(color))
        self.color_preview.setPalette(palette)

    def _test_connection(self):
        self.conn_status.setText("Testing...")
        self.conn_status.setStyleSheet("color: gray")
        QApplication.processEvents()
        
        protocol = ETLProtocol(self.ip_edit.text(), self.port_spin.value())
        info = protocol.get_device_info()
        
        if info:
            self.conn_status.setText(f"✓ {info}")
            self.conn_status.setStyleSheet("color: green")
        else:
            self.conn_status.setText("✗ Connection failed")
            self.conn_status.setStyleSheet("color: red")

    def _auto_detect(self):
        self.conn_status.setText("Detecting...")
        self.conn_status.setStyleSheet("color: gray")
        QApplication.processEvents()
        
        protocol = ETLProtocol(self.ip_edit.text(), self.port_spin.value())
        size = protocol.get_matrix_config()
        
        if size:
            self.inputs_spin.setValue(size[0])
            self.outputs_spin.setValue(size[1])
            self.conn_status.setText(f"✓ Detected {size[0]}×{size[1]}")
            self.conn_status.setStyleSheet("color: green")
        else:
            self.conn_status.setText("✗ Could not detect")
            self.conn_status.setStyleSheet("color: red")

    def _choose_color(self):
        color = QColorDialog.getColor(QColor(self.active_color), self, "Active Route Color")
        if color.isValid():
            self.active_color = color.name()
            self._set_color_preview(self.active_color)

    def _toggle_crosshair_settings(self):
        enabled = self.crosshair_check.isChecked()
        # Brightness is always enabled (affects single-cell hover too)
        # Only crosshair-specific settings are disabled
        self.crosshair_border_btn.setEnabled(enabled)
        self.crosshair_border_preview.setEnabled(enabled)

    def _set_crosshair_border_preview(self, color: str):
        palette = self.crosshair_border_preview.palette()
        palette.setColor(QPalette.Window, QColor(color))
        self.crosshair_border_preview.setPalette(palette)

    def _choose_crosshair_border_color(self):
        color = QColorDialog.getColor(QColor(self.crosshair_border_color), self, "Crosshair Border Color")
        if color.isValid():
            self.crosshair_border_color = color.name()
            self._set_crosshair_border_preview(self.crosshair_border_color)

    def get_values(self) -> dict:
        return {
            'ip_address': self.ip_edit.text().strip(),
            'port': self.port_spin.value(),
            'num_inputs': self.inputs_spin.value(),
            'num_outputs': self.outputs_spin.value(),
            'label_font_family': self.label_font_combo.currentText(),
            'label_font_size': self.label_font_size.value(),
            'button_font_family': self.btn_font_combo.currentText(),
            'button_font_size': self.btn_font_size.value(),
            'active_route_color': self.active_color,
            'dark_theme': self.dark_theme_radio.isChecked(),
            'crosshair_enabled': self.crosshair_check.isChecked(),
            'crosshair_luminance_shift': self.crosshair_lum_spin.value(),
            'crosshair_border_color': self.crosshair_border_color,
            'use_custom_ranges': self.use_custom_check.isChecked(),
            'custom_inputs': parse_range_string(self.custom_inputs_edit.text()),
            'custom_outputs': parse_range_string(self.custom_outputs_edit.text()),
            'show_input_numbers': self.show_input_numbers_check.isChecked(),
            'show_output_numbers': self.show_output_numbers_check.isChecked(),
        }


class SetupWidget(QWidget):
    setup_complete = Signal()

    def __init__(self, config: RouterConfig):
        super().__init__()
        self.config = config
        self.probe_timer = None
        self.last_probed_ip = ''
        self._probed_outputs = 0
        self.additional_router_widgets = []  # List of dicts with ip, port, status widgets
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.addStretch()

        center = QVBoxLayout()
        center.setAlignment(Qt.AlignCenter)

        title = QLabel("Connect to Router")
        title.setFont(QFont("Helvetica", 20, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        center.addWidget(title)

        subtitle = QLabel("Enter the IP address and i/o of the ETL matrix")
        subtitle.setFont(QFont("Helvetica", 11))
        subtitle.setStyleSheet("color: lightgray;")
        subtitle.setAlignment(Qt.AlignCenter)
        center.addWidget(subtitle)
        center.addSpacing(20)

        # Form container with fixed width
        self.form_container = QWidget()
        self.form_container.setFixedWidth(320)
        self.form_layout = QVBoxLayout(self.form_container)
        self.form_layout.setSpacing(10)
        self.form_layout.setContentsMargins(0, 0, 0, 0)

        # IP Address field
        ip_label = QLabel("Router IP Address")
        ip_label.setStyleSheet("font-weight: bold; color: lightgray; font-size: 12px;")
        self.form_layout.addWidget(ip_label)
        
        self.ip_edit = QLineEdit()
        self.ip_edit.setPlaceholderText("e.g. 192.168.1.100")
        self.ip_edit.setMinimumHeight(36)
        self.ip_edit.textChanged.connect(self._on_ip_changed)
        self.form_layout.addWidget(self.ip_edit)
        
        # Status label (below IP field)
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("font-size: 12px;")
        self.form_layout.addWidget(self.status_label)

        # Port field
        port_label = QLabel("Port")
        port_label.setStyleSheet("font-weight: bold; color: lightgray; font-size: 12px;")
        self.form_layout.addWidget(port_label)
        
        self.port_spin = QSpinBox()
        self.port_spin.setRange(1, 65535)
        self.port_spin.setValue(4000)
        self.port_spin.setMinimumHeight(36)
        self.form_layout.addWidget(self.port_spin)

        # Router name field
        name_label = QLabel("Router Name (optional)")
        name_label.setStyleSheet("font-weight: bold; color: lightgray; font-size: 12px;")
        self.form_layout.addWidget(name_label)
        
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("e.g. Golf matrix or Reality Router etc")
        self.name_edit.setMinimumHeight(36)
        self.form_layout.addWidget(self.name_edit)

        # Inputs and Outputs side by side
        io_row = QHBoxLayout()
        io_row.setSpacing(15)
        
        # Inputs
        inputs_col = QVBoxLayout()
        inputs_label = QLabel("Inputs")
        inputs_label.setStyleSheet("font-weight: bold; color: lightgray; font-size: 12px;")
        inputs_col.addWidget(inputs_label)
        self.inputs_spin = QSpinBox()
        self.inputs_spin.setRange(1, 256)
        self.inputs_spin.setValue(8)
        self.inputs_spin.setMinimumHeight(36)
        inputs_col.addWidget(self.inputs_spin)
        io_row.addLayout(inputs_col)
        io_row.setSpacing(10)
        
        # Outputs
        outputs_col = QVBoxLayout()
        outputs_label = QLabel("Outputs")
        outputs_label.setStyleSheet("font-weight: bold; color: lightgray; font-size: 12px;")
        outputs_col.addWidget(outputs_label)
        self.outputs_spin = QSpinBox()
        self.outputs_spin.setRange(1, 256)
        self.outputs_spin.setValue(8)
        self.outputs_spin.setMinimumHeight(36)
        outputs_col.addWidget(self.outputs_spin)
        io_row.addLayout(outputs_col)
        io_row.setSpacing(10)
     
        self.form_layout.addLayout(io_row)
        self.form_layout.addSpacing(10)

        # Advanced section checkbox
        self.advanced_check = QCheckBox("Advanced options")
        self.advanced_check.setStyleSheet("color: lightgray; font-size: 12px;")
        self.advanced_check.stateChanged.connect(self._toggle_advanced)
        self.form_layout.addWidget(self.advanced_check)

        # Advanced container (hidden by default)
        self.advanced_container = QWidget()
        self.advanced_layout = QVBoxLayout(self.advanced_container)
        self.advanced_layout.setContentsMargins(0, 10, 0, 0)
        self.advanced_layout.setSpacing(10)
        
        # Combine routers checkbox
        self.combine_check = QCheckBox("Combine multiple routers")
        self.combine_check.setStyleSheet("color: lightgray; font-size: 12px;")
        self.combine_check.stateChanged.connect(self._toggle_combine)
        self.advanced_layout.addWidget(self.combine_check)
        
        # Additional routers container
        self.additional_routers_container = QWidget()
        self.additional_routers_layout = QVBoxLayout(self.additional_routers_container)
        self.additional_routers_layout.setContentsMargins(0, 0, 0, 0)
        self.additional_routers_layout.setSpacing(10)
        self.additional_routers_container.hide()
        self.advanced_layout.addWidget(self.additional_routers_container)
        
        # Add router button
        self.add_router_btn = QPushButton("+ Add Router")
        self.add_router_btn.setStyleSheet("color: lightgray; font-size: 12px;")
        self.add_router_btn.clicked.connect(self._add_router_row)
        self.add_router_btn.hide()
        self.advanced_layout.addWidget(self.add_router_btn)
        
        self.advanced_container.hide()
        self.form_layout.addWidget(self.advanced_container)

        # Connect button
        self.continue_btn = QPushButton("Connect")
        self.continue_btn.setMinimumHeight(44)
        self.continue_btn.setStyleSheet("""
            QPushButton {
                background-color: #e94560;
                color: white;
                border: none;
                border-radius: 5px;
                font-size: 18px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #d63850;
            }
            QPushButton:pressed {
                background-color: #c42840;
            }
        """)
        self.continue_btn.clicked.connect(self._continue)
        self.form_layout.addWidget(self.continue_btn)

        # Center the form container
        center.addWidget(self.form_container, alignment=Qt.AlignCenter)

        layout.addLayout(center)
        layout.addStretch()

    def _toggle_advanced(self, state):
        """Show/hide advanced options."""
        if state:  # Non-zero means checked
            self.advanced_container.show()
        else:
            self.advanced_container.hide()
            self.combine_check.setChecked(False)

    def _toggle_combine(self, state):
        """Show/hide multi-router options."""
        if state:  # Non-zero means checked
            self.additional_routers_container.show()
            self.add_router_btn.show()
            # Add first additional router if none exist
            if not self.additional_router_widgets:
                self._add_router_row()
        else:
            self.additional_routers_container.hide()
            self.add_router_btn.hide()

    def _add_router_row(self):
        """Add a row for an additional router."""
        row_widget = QWidget()
        row_layout = QVBoxLayout(row_widget)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(5)
        
        # Header with router number and remove button
        header_row = QHBoxLayout()
        router_num = len(self.additional_router_widgets) + 2
        header_label = QLabel(f"Router {router_num}")
        header_label.setStyleSheet("font-weight: bold; color: lightgray; font-size: 12px;")
        header_row.addWidget(header_label)
        header_row.addStretch()
        
        header_port_label = QLabel("Port")
        header_port_label.setStyleSheet("color: lightgray; font-size: 12px;")
        header_row.addWidget(header_port_label)
        header_port_label.setContentsMargins(0, 0, 24, 0)

        remove_btn = QPushButton("×")
        remove_btn.setFixedSize(24, 24)
        remove_btn.setStyleSheet("font-size: 16px; font-weight: bold;")
        remove_btn.clicked.connect(lambda: self._remove_router_row(row_widget))
        header_row.addWidget(remove_btn)
        row_layout.addLayout(header_row)
        
        # IP and Port row
        ip_port_row = QHBoxLayout()
        ip_port_row.setSpacing(10)
        
        ip_edit = QLineEdit()
        ip_edit.setPlaceholderText("IP Address")
        ip_edit.setMinimumHeight(32)
        ip_port_row.addWidget(ip_edit, stretch=3)
        
        port_spin = QSpinBox()
        port_spin.setRange(1, 65535)
        port_spin.setValue(4000)
        # port_spin.setValue(4000 + router_num - 1)
        port_spin.setMinimumHeight(32)
        ip_port_row.addWidget(port_spin, stretch=1)
        
        row_layout.addLayout(ip_port_row)
        
        # Status label
        status_label = QLabel("")
        status_label.setStyleSheet("font-size: 11px;")
        row_layout.addWidget(status_label)
        
        # Store references
        router_data = {
            'widget': row_widget,
            'ip_edit': ip_edit,
            'port_spin': port_spin,
            'status_label': status_label,
            'header_label': header_label,
            'num_inputs': 0,
            'num_outputs': 0
        }
        self.additional_router_widgets.append(router_data)
        
        # Connect probe on IP change
        ip_edit.textChanged.connect(lambda text, rd=router_data: self._probe_additional_router(rd))
        
        self.additional_routers_layout.addWidget(row_widget)

    def _remove_router_row(self, row_widget):
        """Remove an additional router row."""
        for i, data in enumerate(self.additional_router_widgets):
            if data['widget'] == row_widget:
                self.additional_router_widgets.pop(i)
                row_widget.deleteLater()
                break
        
        # Renumber remaining routers
        for i, data in enumerate(self.additional_router_widgets):
            data['header_label'].setText(f"Router {i + 2}")

    def _probe_additional_router(self, router_data):
        """Probe an additional router for connection info."""
        ip = router_data['ip_edit'].text().strip()
        port = router_data['port_spin'].value()
        status_label = router_data['status_label']
        
        if not self._is_valid_ip(ip):
            status_label.setText("")
            router_data['num_inputs'] = 0
            router_data['num_outputs'] = 0
            return
        
        status_label.setText("Detecting...")
        status_label.setStyleSheet("color: gray; font-size: 11px;")
        QApplication.processEvents()
        
        protocol = ETLProtocol(ip, port, timeout=2.0)
        info = protocol.get_device_info()
        
        if not info:
            status_label.setText("No router found")
            status_label.setStyleSheet("color: #e94560; font-size: 11px;")
            router_data['num_inputs'] = 0
            router_data['num_outputs'] = 0
            return
        
        size = protocol.get_matrix_config()
        if size:
            router_data['num_inputs'] = size[0]
            router_data['num_outputs'] = size[1]
            status_label.setText(f"✓ {size[0]}×{size[1]} matrix")
            status_label.setStyleSheet("color: #4ade80; font-size: 11px;")
        else:
            status_label.setText(f"✓ Connected (size unknown)")
            status_label.setStyleSheet("color: orange; font-size: 11px;")

    def _on_ip_changed(self, text):
        """Debounced IP validation - waits 500ms after typing stops."""
        if self.probe_timer:
            self.probe_timer.stop()
        
        if text != self.last_probed_ip:
            self.last_probed_ip = ''
        
        self.probe_timer = QTimer()
        self.probe_timer.setSingleShot(True)
        self.probe_timer.timeout.connect(lambda: self._probe_router(text))
        self.probe_timer.start(500)

    def _is_valid_ip(self, ip: str) -> bool:
        """Check if IP address format is valid."""
        pattern = r'^(\d{1,3}\.){3}\d{1,3}$'
        if not re.match(pattern, ip):
            return False
        octets = ip.split('.')
        return all(0 <= int(octet) <= 255 for octet in octets)

    def _probe_router(self, ip: str):
        """Probe the router to get device info and matrix size."""
        ip = ip.strip()
        
        if not self._is_valid_ip(ip):
            self.last_probed_ip = ''
            self.status_label.setText("")
            return
        
        if ip == self.last_probed_ip:
            return
        
        self.status_label.setText("Detecting router...")
        self.status_label.setStyleSheet("color: gray; font-size: 11px;")
        QApplication.processEvents()
        
        protocol = ETLProtocol(ip, self.port_spin.value(), timeout=2.0)
        
        info = protocol.get_device_info()
        if not info:
            self.last_probed_ip = ''
            self.status_label.setText("No router found")
            self.status_label.setStyleSheet("color: #e94560; font-size: 11px;")
            return
        
        size = protocol.get_matrix_config()
        if size:
            self.last_probed_ip = ip
            self.inputs_spin.setValue(size[0])
            self.outputs_spin.setValue(size[1])
            self._probed_outputs = size[1]  # Store for later
            self.status_label.setText(f"Found: {size[0]}×{size[1]} matrix")
            self.status_label.setStyleSheet("color: #4ade80; font-size: 11px;")
            
            if not self.name_edit.text():
                model = info.split(' - ')[0] if ' - ' in info else info
                self.name_edit.setPlaceholderText(model)
        else:
            self.last_probed_ip = ip
            self.status_label.setText(f"Connected but could not detect size")
            self.status_label.setStyleSheet("color: orange; font-size: 11px;")

    def _continue(self):
        ip = self.ip_edit.text().strip()
        if not ip:
            QMessageBox.critical(self, "Error", "Please enter an IP address.")
            return

        inputs = self.inputs_spin.value()
        outputs = self.outputs_spin.value()  # This is now the TOTAL desired outputs

        self.config.ip_address = ip
        self.config.port = self.port_spin.value()
        self.config.num_inputs = inputs
        self.config.first_run = False
        
        name = self.name_edit.text().strip()
        if name:
            self.config.router_name = name
        else:
            self.config.router_name = ""

        # Handle combined routers
        self.config.combine_routers = self.combine_check.isChecked()
        self.config.additional_routers = []
        
        if self.config.combine_routers:
            # Store primary router's actual output count from probe
            if self._probed_outputs:
                self.config.primary_num_outputs = self._probed_outputs
            else:
                self.config.primary_num_outputs = outputs  # Fallback to spinbox value
            
            # Collect additional routers
            for router_data in self.additional_router_widgets:
                router_ip = router_data['ip_edit'].text().strip()
                if router_ip and self._is_valid_ip(router_ip):
                    self.config.additional_routers.append({
                        'ip': router_ip,
                        'port': router_data['port_spin'].value(),
                        'num_outputs': router_data['num_outputs'] or 16  # Fallback
                    })
            
            # num_outputs is the TOTAL the user wants across all routers
            self.config.num_outputs = outputs
        else:
            # Single router mode
            self.config.primary_num_outputs = outputs
            self.config.num_outputs = outputs

        # Create output groups for the total outputs
        self.config.output_groups = []
        for out in range(1, self.config.num_outputs + 1):
            self.config.output_groups.append(
                OutputGroup(f"Out {out}", "#b0b0b0", [out])
            )

        self.setup_complete.emit()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self._continue()
        else:
            super().keyPressEvent(event)

        

class MatrixButton(QLabel):
    clicked = Signal()
    right_clicked = Signal()
    hover_enter = Signal(int, int)  # input, output
    hover_leave = Signal()

    def __init__(self, text="", min_size=30):
        super().__init__(text)
        self.setAlignment(Qt.AlignCenter)
        self.setCursor(Qt.PointingHandCursor)
        self._min_size = min_size
        self.setMinimumSize(min_size, 20)
        self.setMouseTracking(True)
        self.input_num = 0
        self.output_num = 0
        self._base_style = ""
        self._hover_style = ""
        # Use expanding policy to ensure equal distribution
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

    def set_position(self, input_num: int, output_num: int):
        """Store the button's position in the matrix."""
        self.input_num = input_num
        self.output_num = output_num

    def set_min_width(self, width):
        """Allow dynamic minimum width adjustment."""
        self._min_size = width
        self.setMinimumSize(width, 20)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        elif event.button() == Qt.RightButton:
            self.right_clicked.emit()

    def enterEvent(self, event):
        self.hover_enter.emit(self.input_num, self.output_num)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.hover_leave.emit()
        super().leaveEvent(event)

    def set_color(self, bg_color: str, text_color: str, dark_theme: bool = True, 
                  highlight_right: bool = False, highlight_bottom: bool = False,
                  highlight_border: str = "#ffffff", luminance_shift: int = 0):
        """Set button color with optional border highlighting for crosshair effect."""
        border_color = "#404040" if dark_theme else "#c0c0c0"
        
        # Apply luminance shift if specified
        if luminance_shift != 0:
            bg_color = self._adjust_luminance(bg_color, luminance_shift)
        
        # Determine border colors
        right_border = highlight_border if highlight_right else border_color
        bottom_border = highlight_border if highlight_bottom else border_color
        
        self.setStyleSheet(
            f"background-color: {bg_color}; color: {text_color}; "
            f"border: none; border-right: 1px solid {right_border}; border-bottom: 1px solid {bottom_border};"
        )
    
    def _adjust_luminance(self, hex_color: str, shift: int) -> str:
        """Adjust the luminance of a hex color by a percentage."""
        try:
            hex_color = hex_color.lstrip('#')
            r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            
            # Shift each channel
            factor = 1 + (shift / 100)
            r = max(0, min(255, int(r * factor)))
            g = max(0, min(255, int(g * factor)))
            b = max(0, min(255, int(b * factor)))
            
            return f"#{r:02x}{g:02x}{b:02x}"
        except:
            return hex_color


class MatrixWidget(QWidget):
    def __init__(self, config: RouterConfig, protocol: ETLProtocol):
        super().__init__()
        self.config = config
        self.protocol = protocol
        self.additional_protocols: List[ETLProtocol] = []  # For combined routers
        self.current_routes: Dict[int, int] = {}
        self.route_buttons: Dict[Tuple[int, int], MatrixButton] = {}
        self.input_labels: Dict[int, QLabel] = {}
        self.group_headers: Dict[int, QLabel] = {}
        self.output_to_group: Dict[int, int] = {}
        self.display_groups: List[OutputGroup] = []  # Groups currently displayed
        self.group_select_start: Optional[int] = None
        
        # Crosshair hover tracking
        self.hover_input: Optional[int] = None
        self.hover_output: Optional[int] = None
        self._prev_hover_input: Optional[int] = None
        self._prev_hover_output: Optional[int] = None
        
        # Cache for button states to avoid redundant updates
        self._button_state_cache: Dict[Tuple[int, int], tuple] = {}
        
        # Multi-select support
        self.selected_buttons: set = set()  # Set of (input, output) tuples
        self.multi_select_mode: bool = False
        
        self.signals = RouteSignals()
        self.signals.route_complete.connect(self._on_route_complete)
        
        self.status_callback = None
        self.hint_callback = None
        self.refresh_callback = None  # Called after routing to refresh status
        
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)  # Enable keyboard focus
        self._build_matrix()

    def set_callbacks(self, status_cb, hint_cb, refresh_cb=None):
        self.status_callback = status_cb
        self.hint_callback = hint_cb
        self.refresh_callback = refresh_cb

    def keyPressEvent(self, event):
        """Handle keyboard shortcuts."""
        if event.key() == Qt.Key_Escape:
            # Clear selection
            self._clear_selection()
        elif event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            # Route all selected
            if self.selected_buttons:
                self._route_selected()
        elif event.modifiers() & (Qt.ControlModifier | Qt.ShiftModifier):
            self.multi_select_mode = True
        super().keyPressEvent(event)

    def keyReleaseEvent(self, event):
        """Handle key release."""
        if not (event.modifiers() & (Qt.ControlModifier | Qt.ShiftModifier)):
            self.multi_select_mode = False
        super().keyReleaseEvent(event)

    def _clear_selection(self):
        """Clear all selected buttons."""
        self.selected_buttons.clear()
        self._update_route_display()
        if self.hint_callback:
            self.hint_callback("")

    def _toggle_selection(self, input_num: int, output_num: int):
        """Toggle selection of a button."""
        key = (input_num, output_num)
        if key in self.selected_buttons:
            self.selected_buttons.remove(key)
        else:
            self.selected_buttons.add(key)
        self._update_route_display()
        
        if self.selected_buttons and self.hint_callback:
            self.hint_callback(f"{len(self.selected_buttons)} selected - Press Enter to route, Escape to clear")

    def _route_selected(self):
        """Route all selected buttons."""
        if not self.selected_buttons:
            return
        
        routes_to_make = list(self.selected_buttons)
        self._clear_selection()
        
        if self.status_callback:
            self.status_callback(f"Routing {len(routes_to_make)} crosspoints...")
        
        def do_routes():
            success_count = 0
            for inp, out in routes_to_make:
                protocol, local_output = self._get_protocol_for_output(out)
                if protocol.route(inp, local_output):
                    success_count += 1
                time.sleep(0.1)  # Small delay between routes
            
            # Update UI on completion
            def update():
                if self.status_callback:
                    self.status_callback(f"✓ Routed {success_count}/{len(routes_to_make)} crosspoints")
                if self.refresh_callback:
                    self.refresh_callback()
            
            QTimer.singleShot(0, update)
        
        threading.Thread(target=do_routes, daemon=True).start()

    def leaveEvent(self, event):
        """Clear crosshair when mouse leaves the widget."""
        if self.hover_input is not None or self.hover_output is not None:
            self._prev_hover_input = self.hover_input
            self._prev_hover_output = self.hover_output
            self.hover_input = None
            self.hover_output = None
            self._update_hover_display()
        super().leaveEvent(event)

    def _on_button_hover_enter(self, input_num: int, output_num: int):
        """Handle mouse entering a button - update hover highlight."""
        # Always track hover position (for single-cell highlight or crosshair)
        if self.hover_input != input_num or self.hover_output != output_num:
            self._prev_hover_input = self.hover_input
            self._prev_hover_output = self.hover_output
            self.hover_input = input_num
            self.hover_output = output_num
            self._update_hover_display()

    def _on_button_hover_leave(self):
        """Handle mouse leaving a button."""
        # Don't clear immediately - let enterEvent of next button or leaveEvent of widget handle it
        pass

    def _update_hover_display(self):
        """Optimized hover update - only updates buttons affected by hover change."""
        crosshair_enabled = self.config.crosshair_enabled
        
        # Collect buttons that need updating
        buttons_to_update = set()
        
        if crosshair_enabled:
            # For crosshair mode, update entire previous and current rows/columns
            all_inputs = self.config.get_inputs()
            all_outputs = []
            for group in self.display_groups:
                all_outputs.extend(group.outputs)
            
            # Previous crosshair buttons
            if self._prev_hover_input is not None:
                for out in all_outputs:
                    buttons_to_update.add((self._prev_hover_input, out))
                try:
                    prev_idx = all_inputs.index(self._prev_hover_input)
                    if prev_idx > 0:
                        for out in all_outputs:
                            buttons_to_update.add((all_inputs[prev_idx - 1], out))
                except ValueError:
                    pass
                    
            if self._prev_hover_output is not None:
                for inp in all_inputs:
                    buttons_to_update.add((inp, self._prev_hover_output))
                try:
                    prev_idx = all_outputs.index(self._prev_hover_output)
                    if prev_idx > 0:
                        for inp in all_inputs:
                            buttons_to_update.add((inp, all_outputs[prev_idx - 1]))
                except ValueError:
                    pass
            
            # Current crosshair buttons
            if self.hover_input is not None:
                for out in all_outputs:
                    buttons_to_update.add((self.hover_input, out))
                try:
                    curr_idx = all_inputs.index(self.hover_input)
                    if curr_idx > 0:
                        for out in all_outputs:
                            buttons_to_update.add((all_inputs[curr_idx - 1], out))
                except ValueError:
                    pass
                    
            if self.hover_output is not None:
                for inp in all_inputs:
                    buttons_to_update.add((inp, self.hover_output))
                try:
                    curr_idx = all_outputs.index(self.hover_output)
                    if curr_idx > 0:
                        for inp in all_inputs:
                            buttons_to_update.add((inp, all_outputs[curr_idx - 1]))
                except ValueError:
                    pass
        else:
            # For single-cell mode, only update previous and current cell
            if self._prev_hover_input is not None and self._prev_hover_output is not None:
                buttons_to_update.add((self._prev_hover_input, self._prev_hover_output))
            if self.hover_input is not None and self.hover_output is not None:
                buttons_to_update.add((self.hover_input, self.hover_output))
        
        # Update only the affected buttons
        self._update_buttons(buttons_to_update)

    def _update_buttons(self, buttons_to_update: set):
        """Update only the specified buttons."""
        if not buttons_to_update:
            return
            
        active_color = self.config.active_route_color
        dark_theme = self.config.dark_theme
        crosshair_enabled = self.config.crosshair_enabled
        hover_lum = self.config.crosshair_luminance_shift
        crosshair_border = self.config.crosshair_border_color
        selection_border = "#ffff00"

        # Pre-compute indices
        all_outputs = []
        for group in self.display_groups:
            all_outputs.extend(group.outputs)
        output_to_idx = {out: idx for idx, out in enumerate(all_outputs)}
        
        all_inputs = self.config.get_inputs()
        input_to_idx = {inp: idx for idx, inp in enumerate(all_inputs)}
        
        hover_out_idx = output_to_idx.get(self.hover_output, -1) if crosshair_enabled else -1
        hover_inp_idx = input_to_idx.get(self.hover_input, -1) if crosshair_enabled else -1

        for inp, out in buttons_to_update:
            btn = self.route_buttons.get((inp, out))
            if btn is None:
                continue
                
            base_color = self._get_output_color(out)
            out_idx = output_to_idx.get(out, -1)
            inp_idx = input_to_idx.get(inp, -1)
            
            is_hovered_cell = (inp == self.hover_input and out == self.hover_output)
            in_row = hover_inp_idx >= 0 and inp_idx == hover_inp_idx
            in_col = hover_out_idx >= 0 and out_idx == hover_out_idx
            in_crosshair = in_row or in_col
            
            highlight_right = (out_idx == hover_out_idx or out_idx == hover_out_idx - 1) if hover_out_idx >= 0 else False
            highlight_bottom = (inp_idx == hover_inp_idx or inp_idx == hover_inp_idx - 1) if hover_inp_idx >= 0 else False
            
            row_lum = self.config.row_luminance.get(inp, 0)
            is_selected = (inp, out) in self.selected_buttons
            
            if self.current_routes.get(out) == inp:
                color = active_color
            else:
                color = base_color
            
            text_color = self._get_contrast_color(color)
            
            total_lum_shift = row_lum
            if in_crosshair or is_hovered_cell:
                total_lum_shift += hover_lum
            
            if is_selected:
                btn.set_color(color, text_color, dark_theme, 
                             highlight_right=True, highlight_bottom=True,
                             highlight_border=selection_border,
                             luminance_shift=total_lum_shift + 15)
            else:
                btn.set_color(color, text_color, dark_theme, 
                             highlight_right=highlight_right, 
                             highlight_bottom=highlight_bottom,
                             highlight_border=crosshair_border,
                             luminance_shift=total_lum_shift)

    def _get_contrast_color(self, hex_color: str) -> str:
        try:
            hex_color = hex_color.lstrip('#')
            r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
            return "#000000" if luminance > 0.5 else "#ffffff"
        except:
            return "#000000"

    def _build_output_to_group_map(self):
        """Build mapping from output number to display group index."""
        self.output_to_group.clear()
        for idx, group in enumerate(self.display_groups):
            for out in group.outputs:
                self.output_to_group[out] = idx

    def _get_output_color(self, output: int) -> str:
        if output in self.output_to_group:
            group_idx = self.output_to_group[output]
            if group_idx < len(self.display_groups):
                return self.display_groups[group_idx].color
        return "#b0b0b0"

    def _build_matrix(self):
        if self.layout():
            QWidget().setLayout(self.layout())
        
        self.route_buttons.clear()
        self.input_labels.clear()
        self.group_headers.clear()

        layout = QGridLayout(self)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)

        # Get actual input/output lists
        inputs = self.config.get_inputs()
        outputs = self.config.get_outputs()
        
        # Get display groups (filtered for visible outputs)
        self.display_groups = self.config.get_display_groups()
        self._build_output_to_group_map()

        if not inputs or not outputs:
            layout.addWidget(QLabel("Please configure router settings"), 0, 0)
            return

        # Calculate button minimum size based on screen width
        screen = QApplication.primaryScreen().geometry()
        available_width = screen.width() - 150
        min_btn_width = max(15, min(30, available_width // len(outputs)))
        
        dark_theme = self.config.dark_theme
        compact_mode = self.config.compact_mode
        border_color = "#404040" if dark_theme else "#c0c0c0"
        label_bg = "#535353" if dark_theme else "#f0f0f0"
        label_text = "white" if dark_theme else "black"
        input_bg = "#808080" if dark_theme else "#e0e0e0"

        label_font = QFont(self.config.label_font_family, self.config.label_font_size, QFont.Bold)
        button_font = QFont(self.config.button_font_family, self.config.button_font_size)

        if compact_mode:
            min_btn_width = max(10, min(20, available_width // len(outputs)))
        
        start_row = 0
        start_col = 0
        
        if not compact_mode:
            # Corner cell
            corner = QLabel("")
            corner.setStyleSheet(f"background-color: {label_bg}; border-right: 1px solid {border_color}; border-bottom: 1px solid {border_color};")
            corner.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
            layout.addWidget(corner, 0, 0)
            start_col = 1
            start_row = 1
            
            # Input number column header (if enabled)
            if self.config.show_input_numbers:
                num_header = QLabel("#")
                num_header.setFont(label_font)
                num_header.setAlignment(Qt.AlignCenter)
                num_header.setStyleSheet(
                    f"background-color: {label_bg}; color: {label_text}; padding: 4px; "
                    f"border-right: 1px solid {border_color}; border-bottom: 1px solid {border_color};"
                )
                layout.addWidget(num_header, 0, 1)
                start_col = 2

            # Column headers using display groups
            col = 2 if self.config.show_input_numbers else 1
            for group_idx, group in enumerate(self.display_groups):
                span = len(group.outputs)

                if self.config.show_output_numbers:
                    if len(group.outputs) == 1:
                        header_text = f"{group.outputs[0]}: {group.name}"
                    else:
                        header_text = f"{min(group.outputs)}-{max(group.outputs)}: {group.name}"
                else:
                    header_text = group.name

                if self.config.show_output_numbers:
                    if len(group.outputs) == 1:
                        header_text = f"{group.outputs[0]}: {group.name}"
                    else:
                        header_text = f"{min(group.outputs)}-{max(group.outputs)}: {group.name}"
                else:
                    header_text = group.name
                header = QLabel(header_text)
                header.setFont(label_font)
                header.setAlignment(Qt.AlignCenter)
                header.setCursor(Qt.PointingHandCursor)
                header.setMinimumWidth(0)
                header.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
                text_color = self._get_contrast_color(group.color)
                header.setStyleSheet(
                    f"background-color: {group.color}; color: {text_color}; padding: 4px; "
                    f"border-right: 1px solid {border_color}; border-bottom: 1px solid {border_color};"
                )
                header.mousePressEvent = lambda e, idx=group_idx: self._on_header_click(e, idx)
                layout.addWidget(header, 0, col, 1, span)
                self.group_headers[group_idx] = header
                col += span

        # Input rows
        for row_idx, inp in enumerate(inputs):
            row = row_idx + start_row
            
            if not compact_mode:
                # Input label
                name = self.config.input_names.get(inp, f"Input {inp}")
                label = QLabel(name)
                label.setFont(label_font)
                label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
                label.setStyleSheet(
                    f"background-color: {input_bg}; color: {label_text}; padding: 4px; "
                    f"border-right: 1px solid {border_color}; border-bottom: 1px solid {border_color};"
                )
                label.setCursor(Qt.PointingHandCursor)
                label.mousePressEvent = lambda e, i=inp: self._on_input_click(e, i)
                layout.addWidget(label, row, 0)
                self.input_labels[inp] = label
                
                # Input number column (if enabled)
                if self.config.show_input_numbers:
                    num_label = QLabel(str(inp))
                    num_label.setFont(label_font)
                    num_label.setAlignment(Qt.AlignCenter)
                    num_label.setStyleSheet(
                        f"background-color: {input_bg}; color: {label_text}; padding: 4px; "
                        f"border-right: 1px solid {border_color}; border-bottom: 1px solid {border_color};"
                    )
                    layout.addWidget(num_label, row, 1)

            col = start_col
            for group in self.display_groups:
                for out in group.outputs:
                    btn_label = "" if compact_mode else self.config.button_labels.get(str(inp), "○")
                    btn = MatrixButton(btn_label, min_size=min_btn_width)
                    btn.setFont(button_font)
                    btn.set_position(inp, out)
                    if compact_mode:
                        btn.setToolTip(f"Input {inp} → Output {out}")
                    btn_color = group.color
                    text_color = self._get_contrast_color(btn_color)
                    btn.set_color(btn_color, text_color, dark_theme)
                    btn.clicked.connect(lambda i=inp, o=out: self._route(i, o))
                    btn.right_clicked.connect(lambda i=inp, o=out: self._button_context_menu(i, o))
                    btn.hover_enter.connect(self._on_button_hover_enter)
                    btn.hover_leave.connect(self._on_button_hover_leave)
                    layout.addWidget(btn, row, col)
                    self.route_buttons[(inp, out)] = btn
                    col += 1

        if not compact_mode:
            layout.setColumnStretch(0, 0)
            if self.config.show_input_numbers:
                layout.setColumnStretch(1, 0)
        for c in range(start_col, len(outputs) + start_col):
            layout.setColumnStretch(c, 1)

    def _find_main_group_index(self, display_group: OutputGroup) -> Optional[int]:
        """Find the index of a display group in the main output_groups list."""
        for idx, main_group in enumerate(self.config.output_groups):
            if set(display_group.outputs) & set(main_group.outputs):
                return idx
        return None

    def _on_header_click(self, event, display_group_idx: int):
        if event.button() == Qt.RightButton:
            self._group_context_menu(display_group_idx)
            return

        if self.group_select_start is None:
            self.group_select_start = display_group_idx
            group = self.display_groups[display_group_idx]
            if self.hint_callback:
                self.hint_callback(f"Click another output to group with '{group.name}'")
        else:
            start_idx = self.group_select_start
            end_idx = display_group_idx
            self.group_select_start = None
            if self.hint_callback:
                self.hint_callback("")

            if start_idx != end_idx:
                self._create_group_from_display_range(start_idx, end_idx)

    def _create_group_from_display_range(self, start_idx: int, end_idx: int):
        """Create a group from a range of display groups."""
        if start_idx > end_idx:
            start_idx, end_idx = end_idx, start_idx

        # Collect all outputs from the display groups in this range
        merged_outputs = []
        for idx in range(start_idx, end_idx + 1):
            merged_outputs.extend(self.display_groups[idx].outputs)
        merged_outputs = sorted(set(merged_outputs))

        from PySide6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(self, "Create Group",
            f"Creating group with outputs {merged_outputs}\nEnter group name:",
            text=f"Group {merged_outputs[0]}-{merged_outputs[-1]}")

        if ok and name:
            color = QColorDialog.getColor(QColor("#4a90d9"), self, "Group Colour")
            if color.isValid():
                # Remove these outputs from any existing groups in main config
                for group in self.config.output_groups[:]:
                    group.outputs = [o for o in group.outputs if o not in merged_outputs]
                
                # Remove empty groups
                self.config.output_groups = [g for g in self.config.output_groups if g.outputs]
                
                # Add new group
                new_group = OutputGroup(name, color.name(), merged_outputs)
                
                # Insert at appropriate position (sorted by first output)
                insert_idx = 0
                for i, g in enumerate(self.config.output_groups):
                    if g.outputs and min(g.outputs) < min(merged_outputs):
                        insert_idx = i + 1
                
                self.config.output_groups.insert(insert_idx, new_group)
                self._build_matrix()
                self._update_route_display()

    def _group_context_menu(self, display_group_idx: int):
        from PySide6.QtWidgets import QMenu, QInputDialog
        self.group_select_start = None
        if self.hint_callback:
            self.hint_callback("")

        if display_group_idx >= len(self.display_groups):
            return
            
        display_group = self.display_groups[display_group_idx]
        
        menu = QMenu(self)
        
        rename_action = menu.addAction(f"Rename '{display_group.name}'")
        color_action = menu.addAction("Change Colour")
        
        ungroup_action = None
        if len(display_group.outputs) > 1:
            menu.addSeparator()
            ungroup_action = menu.addAction("Ungroup (split to individual outputs)")

        action = menu.exec_(self.cursor().pos())
        
        if action == rename_action:
            new_name, ok = QInputDialog.getText(self, "Rename Group", "Enter new name:", text=display_group.name)
            if ok and new_name:
                # Update in main config
                for main_group in self.config.output_groups:
                    if set(display_group.outputs) <= set(main_group.outputs):
                        main_group.name = new_name
                        break
                display_group.name = new_name
                if display_group_idx in self.group_headers:
                    self.group_headers[display_group_idx].setText(new_name)
                    
        elif action == color_action:
            color = QColorDialog.getColor(QColor(display_group.color), self, "Group Colour")
            if color.isValid():
                # Update in main config
                for main_group in self.config.output_groups:
                    if set(display_group.outputs) & set(main_group.outputs):
                        main_group.color = color.name()
                self._build_matrix()
                self._update_route_display()
                
        elif action == ungroup_action:
            outputs_to_ungroup = sorted(display_group.outputs)
            
            # Remove these outputs from existing groups in main config
            for group in self.config.output_groups[:]:
                group.outputs = [o for o in group.outputs if o not in outputs_to_ungroup]
            
            # Remove empty groups
            self.config.output_groups = [g for g in self.config.output_groups if g.outputs]
            
            # Add individual groups for each output
            for out in outputs_to_ungroup:
                new_group = OutputGroup(f"Out {out}", "#b0b0b0", [out])
                # Insert at appropriate position
                insert_idx = 0
                for i, g in enumerate(self.config.output_groups):
                    if g.outputs and min(g.outputs) < out:
                        insert_idx = i + 1
                self.config.output_groups.insert(insert_idx, new_group)
            
            self._build_matrix()
            self._update_route_display()

    def _on_input_click(self, event, inp: int):
        from PySide6.QtWidgets import QMenu, QInputDialog
        
        if event.button() == Qt.RightButton or event.button() == Qt.LeftButton:
            menu = QMenu(self)
            rename_label = menu.addAction("Rename Input Label")
            rename_btn = menu.addAction("Rename Button Labels")
            menu.addSeparator()
            
            # Row luminance submenu
            lum_menu = menu.addMenu("Row Brightness")
            current_lum = self.config.row_luminance.get(inp, 0)
            
            lum_darker = lum_menu.addAction("Darker (-20%)")
            lum_dark = lum_menu.addAction("Slightly Darker (-10%)")
            lum_normal = lum_menu.addAction("Normal (0%)")
            lum_light = lum_menu.addAction("Slightly Brighter (+10%)")
            lum_lighter = lum_menu.addAction("Brighter (+20%)")
            lum_menu.addSeparator()
            lum_custom = lum_menu.addAction(f"Custom... (current: {current_lum}%)")
            lum_menu.addSeparator()
            lum_clear_all = lum_menu.addAction(f"Clear All ({len(self.config.row_luminance)} rows adjusted)")
            lum_clear_all.setEnabled(len(self.config.row_luminance) > 0)
            
            # Mark current setting
            if current_lum == -20:
                lum_darker.setText("✓ Darker (-20%)")
            elif current_lum == -10:
                lum_dark.setText("✓ Slightly Darker (-10%)")
            elif current_lum == 0:
                lum_normal.setText("✓ Normal (0%)")
            elif current_lum == 10:
                lum_light.setText("✓ Slightly Brighter (+10%)")
            elif current_lum == 20:
                lum_lighter.setText("✓ Brighter (+20%)")
            
            action = menu.exec_(self.cursor().pos())
            
            if action == rename_label:
                current = self.config.input_names.get(inp, f"Input {inp}")
                new_name, ok = QInputDialog.getText(self, "Rename Input", f"Name for Input {inp}:", text=current)
                if ok and new_name:
                    self.config.input_names[inp] = new_name
                    self.input_labels[inp].setText(new_name)
            elif action == rename_btn:
                current = self.config.button_labels.get(str(inp), "")
                new_label, ok = QInputDialog.getText(self, "Rename Button",
                    f"Enter button label for Input {inp} row\n(leave empty for default '○'):", text=current)
                if ok:
                    if new_label:
                        self.config.button_labels[str(inp)] = new_label
                    else:
                        self.config.button_labels.pop(str(inp), None)
                    self._update_route_display()
            elif action == lum_darker:
                self._set_row_luminance(inp, -20)
            elif action == lum_dark:
                self._set_row_luminance(inp, -10)
            elif action == lum_normal:
                self._set_row_luminance(inp, 0)
            elif action == lum_light:
                self._set_row_luminance(inp, 10)
            elif action == lum_lighter:
                self._set_row_luminance(inp, 20)
            elif action == lum_custom:
                value, ok = QInputDialog.getInt(self, "Row Brightness",
                    f"Enter brightness adjustment for Input {inp} row\n(-50 to +50 percent):",
                    current_lum, -50, 50)
                if ok:
                    self._set_row_luminance(inp, value)
            elif action == lum_clear_all:
                self._clear_all_row_luminance()
    
    def _clear_all_row_luminance(self):
        """Clear all row brightness adjustments."""
        count = len(self.config.row_luminance)
        self.config.row_luminance.clear()
        self._update_route_display()
        if self.status_callback:
            self.status_callback(f"Cleared brightness adjustments for {count} rows")
    
    def _set_row_luminance(self, inp: int, value: int):
        """Set luminance adjustment for a row."""
        if value == 0:
            self.config.row_luminance.pop(inp, None)
        else:
            self.config.row_luminance[inp] = value
        self._update_route_display()

    def _button_context_menu(self, inp: int, out: int):
        from PySide6.QtWidgets import QMenu, QInputDialog
        
        menu = QMenu(self)
        rename_btn = menu.addAction("Rename Button Labels")
        menu.addSeparator()
        route_action = menu.addAction(f"Route Input {inp} → Output {out}")
        
        action = menu.exec_(self.cursor().pos())
        
        if action == rename_btn:
            current = self.config.button_labels.get(str(inp), "")
            new_label, ok = QInputDialog.getText(self, "Rename Button",
                f"Enter button label for Input {inp} row\n(leave empty for default '○'):", text=current)
            if ok:
                if new_label:
                    self.config.button_labels[str(inp)] = new_label
                else:
                    self.config.button_labels.pop(str(inp), None)
                self._update_route_display()
        elif action == route_action:
            self._route(inp, out)

    def _get_protocol_for_output(self, output: int) -> Tuple[ETLProtocol, int]:
        """Get the correct protocol and local output number for a combined output."""
        router_idx, local_output = self.config.get_router_for_output(output)
        
        if router_idx == 0:
            return (self.protocol, local_output)
        elif router_idx - 1 < len(self.additional_protocols):
            return (self.additional_protocols[router_idx - 1], local_output)
        else:
            # Fallback to primary
            return (self.protocol, output)

    def _route(self, input_num: int, output_num: int):
        # Check if Ctrl or Shift is held for multi-select
        modifiers = QApplication.keyboardModifiers()
        if modifiers & (Qt.ControlModifier | Qt.ShiftModifier):
            self._toggle_selection(input_num, output_num)
            return
        
        # Get the correct protocol for this output (handles combined routers)
        protocol, local_output = self._get_protocol_for_output(output_num)
        
        # Normal single route
        if self.status_callback:
            router_idx, _ = self.config.get_router_for_output(output_num)
            if router_idx > 0:
                self.status_callback(f"Routing Input {input_num} to Output {output_num} (Router {router_idx + 1})...")
            else:
                self.status_callback(f"Routing Input {input_num} to Output {output_num}...")

        def do_route():
            success = protocol.route(input_num, local_output)
            self.signals.route_complete.emit(input_num, output_num, success)

        threading.Thread(target=do_route, daemon=True).start()

    def _on_route_complete(self, input_num: int, output_num: int, success: bool):
        if success:
            if self.status_callback:
                self.status_callback(f"✓ Routed Input {input_num} → Output {output_num}")
            # Update local display immediately
            self.current_routes[output_num] = input_num
            self._update_route_display()
            # Trigger refresh to verify route actually took effect
            if self.refresh_callback:
                QTimer.singleShot(500, self.refresh_callback)
        else:
            if self.status_callback:
                self.status_callback("✗ Failed to route")
            QMessageBox.critical(self, "Route Failed", "Check connection to router.")

    def _update_route_display(self):
        active_color = self.config.active_route_color
        dark_theme = self.config.dark_theme
        crosshair_enabled = self.config.crosshair_enabled
        hover_lum = self.config.crosshair_luminance_shift
        crosshair_border = self.config.crosshair_border_color
        compact_mode = self.config.compact_mode
        selection_border = "#ffff00"

        # Pre-compute output and input index mappings for O(1) lookup
        all_outputs = []
        for group in self.display_groups:
            all_outputs.extend(group.outputs)
        output_to_idx = {out: idx for idx, out in enumerate(all_outputs)}
        
        all_inputs = self.config.get_inputs()
        input_to_idx = {inp: idx for idx, inp in enumerate(all_inputs)}
        
        hover_out_idx = output_to_idx.get(self.hover_output, -1) if crosshair_enabled else -1
        hover_inp_idx = input_to_idx.get(self.hover_input, -1) if crosshair_enabled else -1
        
        # Pre-compute output to group info for tooltips
        output_group_info = {}
        for group in self.display_groups:
            for i, out in enumerate(group.outputs):
                if len(group.outputs) > 1:
                    output_group_info[out] = (group.name, i + 1)
                else:
                    output_group_info[out] = (f"Output {out}", 1)

        for (inp, out), btn in self.route_buttons.items():
            base_color = self._get_output_color(out)
            
            out_idx = output_to_idx.get(out, -1)
            inp_idx = input_to_idx.get(inp, -1)
            
            is_hovered_cell = (inp == self.hover_input and out == self.hover_output)
            
            in_row = hover_inp_idx >= 0 and inp_idx == hover_inp_idx
            in_col = hover_out_idx >= 0 and out_idx == hover_out_idx
            in_crosshair = in_row or in_col
            
            highlight_right = (out_idx == hover_out_idx or out_idx == hover_out_idx - 1) if hover_out_idx >= 0 else False
            highlight_bottom = (inp_idx == hover_inp_idx or inp_idx == hover_inp_idx - 1) if hover_inp_idx >= 0 else False
            
            row_lum = self.config.row_luminance.get(inp, 0)
            is_selected = (inp, out) in self.selected_buttons
            
            if compact_mode:
                btn_label = ""
                group_name, col_idx = output_group_info.get(out, (f"Output {out}", 1))
                input_name = self.config.input_names.get(inp, f"Input {inp}")
                btn.setToolTip(f"{input_name} → {group_name} ({col_idx})")
            else:
                btn_label = self.config.button_labels.get(str(inp), "○")
                btn.setToolTip("")
            
            if self.current_routes.get(out) == inp:
                color = active_color
            else:
                color = base_color
            
            text_color = self._get_contrast_color(color)
            btn.setText(btn_label)
            
            total_lum_shift = row_lum
            if in_crosshair or is_hovered_cell:
                total_lum_shift += hover_lum
            
            if is_selected:
                btn.set_color(color, text_color, dark_theme, 
                             highlight_right=True, highlight_bottom=True,
                             highlight_border=selection_border,
                             luminance_shift=total_lum_shift + 20)
            else:
                btn.set_color(color, text_color, dark_theme, 
                             highlight_right=highlight_right, 
                             highlight_bottom=highlight_bottom,
                             highlight_border=crosshair_border,
                             luminance_shift=total_lum_shift)
    
    def update_routes_from_telemetry(self, routes: dict):
        self.current_routes = routes
        self._update_route_display()

    def rebuild(self):
        self.hover_input = None
        self.hover_output = None
        self._prev_hover_input = None
        self._prev_hover_output = None
        self._button_state_cache.clear()
        self._build_matrix()
        self._update_route_display()


class MainWindow(QMainWindow):
    # Signals for thread-safe UI updates
    connection_status_changed = Signal(bool)
    refresh_complete = Signal(object, bool)  # routes (dict), silent
    refresh_error = Signal(str, bool)  # error message, silent
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ETL RF Matrix Controller")

        # Set window icon (for taskbar on Windows)
        import platform
        if platform.system() == 'Windows':
            from PySide6.QtGui import QIcon
            icon_path = os.path.join(os.path.dirname(__file__), 'icon_1024.ico')
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        
        # Connect signals to slots
        self.connection_status_changed.connect(self._apply_connection_indicator)
        self.refresh_complete.connect(self._on_refresh_complete)
        self.refresh_error.connect(self._on_refresh_error)
        
        self.config = RouterConfig()
        self.config_file = self._get_config_path()
        self._load_config()
        
        self.protocol = None
        self.additional_protocols = []
        self.matrix_widget = None
        self.telemetry_window = None
        self.bg_poll_thread = None
        self.bg_poll_timer = None
        self.toolbar = None
        self.toolbar_buttons = {}
        
        if self.config.first_run or not self.config.ip_address:
            self._show_setup()
        else:
            self._show_main()

    def _get_config_path(self) -> str:
        """Get the path for the config file in a user-writable location."""
        import platform
        
        if platform.system() == "Darwin":  # macOS
            config_dir = os.path.expanduser("~/Library/Application Support/ETL RF Matrix Controller")
        elif platform.system() == "Windows":
            config_dir = os.path.join(os.environ.get("APPDATA", ""), "ETL RF Matrix Controller")
        else:  # Linux and others
            config_dir = os.path.expanduser("~/.config/etl-rf-matrix-controller")
        
        os.makedirs(config_dir, exist_ok=True)
        return os.path.join(config_dir, "etl_config.json")

    def _show_setup(self):
        self.resize(350, 450)
        setup = SetupWidget(self.config)
        setup.setup_complete.connect(self._on_setup_complete)
        self.setCentralWidget(setup)

    def _on_setup_complete(self):
        self._save_config()
        self._show_main()

    def _show_main(self):
        self.protocol = ETLProtocol(self.config.ip_address, self.config.port)

        # Create additional protocols for combined routers
        self.additional_protocols = []
        if self.config.combine_routers and self.config.additional_routers:
            for router in self.config.additional_routers:
                protocol = ETLProtocol(router['ip'], router['port'])
                self.additional_protocols.append(protocol)
        
        # Update window title with router name
        if self.config.router_name:
            self.setWindowTitle(f"ETL RF Matrix Controller — {self.config.router_name}")
        else:
            self.setWindowTitle("ETL RF Matrix Controller")
        
        # Calculate size based on actual displayed inputs/outputs
        inputs = self.config.get_inputs()
        outputs = self.config.get_outputs()
        
        # Get screen dimensions
        screen = QApplication.primaryScreen().geometry()
        screen_width = int(screen.width() * 0.95)
        screen_height = int(screen.height() * 0.85)
        
        # Calculate ideal size
        w = max(800, 100 + len(outputs) * 35)
        h = max(500, 100 + len(inputs) * 25)
        
        if w > screen_width:
            w = screen_width
        
        w = min(w, screen_width)
        h = min(h, screen_height)
        self.resize(w, h)

        self._create_menu()
        
        self.toolbar = QToolBar()
        self.toolbar.setMovable(False)
        self.addToolBar(self.toolbar)
        
        self.toolbar_buttons = {}
        self.toolbar_button_widgets = {}

        # Settings Button
        settings_btn = QPushButton("⚙️ Settings")
        settings_btn.clicked.connect(self._show_settings)
        settings_action = self.toolbar.addWidget(settings_btn)
        self.toolbar_buttons['settings'] = settings_action
        self.toolbar_button_widgets['settings'] = settings_btn
        
        # Refresh Button
        refresh_btn = QPushButton("🔄 Refresh")
        refresh_btn.clicked.connect(self._refresh_status)
        refresh_action = self.toolbar.addWidget(refresh_btn)
        self.toolbar_buttons['refresh'] = refresh_action
        self.toolbar_button_widgets['refresh'] = refresh_btn
        
        # Telemetry Button
        telemetry_btn = QPushButton("📊 Telemetry")
        telemetry_btn.clicked.connect(self._show_telemetry)
        telemetry_action = self.toolbar.addWidget(telemetry_btn)
        self.toolbar_buttons['telemetry'] = telemetry_action
        self.toolbar_button_widgets['telemetry'] = telemetry_btn
                
        # Presets button
        presets_btn = QPushButton("📋 Presets")
        presets_btn.clicked.connect(self._show_presets_menu)
        presets_action = self.toolbar.addWidget(presets_btn)
        self.toolbar_buttons['presets'] = presets_action
        self.toolbar_button_widgets['presets'] = presets_btn
        
        # Compact mode toggle
        self.compact_btn = QPushButton("▫ Compact")
        self.compact_btn.setCheckable(True)
        self.compact_btn.setChecked(self.config.compact_mode)
        self.compact_btn.clicked.connect(self._toggle_compact_mode)
        compact_action = self.toolbar.addWidget(self.compact_btn)
        self.toolbar_buttons['compact'] = compact_action
        self.toolbar_button_widgets['compact'] = self.compact_btn
        
        # Fit to screen button
        fit_btn = QPushButton("💢 Fit")
        fit_btn.setToolTip("Shrink window to fit screen")
        fit_btn.clicked.connect(self._fit_to_screen)
        fit_action = self.toolbar.addWidget(fit_btn)
        self.toolbar_buttons['fit'] = fit_action
        self.toolbar_button_widgets['fit'] = fit_btn
        
        self.hint_label = QLabel("")
        self.hint_label.setStyleSheet("padding: 0 10px;")
        self.toolbar.addWidget(self.hint_label)
        
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.toolbar.addWidget(spacer)
        
        # Connection status indicator
        self.conn_widget = QWidget()
        conn_layout = QHBoxLayout(self.conn_widget)
        conn_layout.setContentsMargins(0, 0, 0, 0)
        conn_layout.setSpacing(2)
        
        self.conn_status_indicator = QLabel("●")
        self.conn_status_indicator.setStyleSheet("color: gray; font-size: 16px;")
        self.conn_status_indicator.setToolTip("Connection status unknown")
        conn_layout.addWidget(self.conn_status_indicator)
        
        self.conn_label = QLabel(f" {self.config.ip_address}")
        conn_layout.addWidget(self.conn_label)
        
        conn_action = self.toolbar.addWidget(self.conn_widget)
        self.toolbar_buttons['connection'] = conn_action
        self.toolbar_button_widgets['connection'] = self.conn_widget
        
        self._apply_toolbar_visibility()

        self.matrix_widget = MatrixWidget(self.config, self.protocol)
        self.matrix_widget.additional_protocols = self.additional_protocols
        self.matrix_widget.set_callbacks(self._set_status, self._set_hint, lambda: self._trigger_refresh())
        self.setCentralWidget(self.matrix_widget)

        self.statusBar().showMessage(f"Connected to {self.config.ip_address}")
        
        self._apply_theme()
        
        # Start connection status checker
        self.conn_check_timer = QTimer(self)
        self.conn_check_timer.timeout.connect(self._check_connection_status)
        self.conn_check_timer.start(10000)
        QTimer.singleShot(100, self._check_connection_status)

        # Background polling for route status
        if self.config.combine_routers and self.additional_protocols:
            # Use custom multi-router polling via QTimer
            self.bg_poll_thread = None
            self.bg_poll_timer = QTimer(self)
            self.bg_poll_timer.timeout.connect(self._poll_all_routers)
            self.bg_poll_timer.start(5000)
            QTimer.singleShot(500, self._poll_all_routers)
        else:
            # Single router - use existing thread
            self.bg_poll_timer = None
            self.bg_poll_thread = TelemetryThread(self.protocol, interval=5.0)
            self.bg_poll_thread.poll_matrix = False
            self.bg_poll_thread.poll_chassis = False
            self.bg_poll_thread.signals.status_received.connect(self.matrix_widget.update_routes_from_telemetry)
            self.bg_poll_thread.start()

    def _trigger_refresh(self):
        """Trigger an immediate status refresh."""
        pass

    def _create_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("File")
        file_menu.addAction("Save Configuration", self._save_config)
        file_menu.addAction("Load Configuration", self._load_config_file)
        file_menu.addAction("Export Config As...", self._export_config)
        file_menu.addSeparator()
        file_menu.addAction("Settings", self._show_settings)
        file_menu.addSeparator()
        file_menu.addAction("Reset to Defaults", self._reset_to_defaults)
        file_menu.addSeparator()
        file_menu.addAction("Exit", self.close)

        view_menu = menubar.addMenu("View")
        view_menu.addAction("Refresh Status", self._refresh_status)
        view_menu.addAction("Fit Window to Screen", self._fit_to_screen)
        view_menu.addSeparator()
        
        self.toolbar_action = view_menu.addAction("Show Toolbar")
        self.toolbar_action.setCheckable(True)
        self.toolbar_action.setChecked(self.config.show_toolbar)
        self.toolbar_action.triggered.connect(self._toggle_toolbar)
        
        toolbar_buttons_menu = view_menu.addMenu("Toolbar Buttons")
        self.toolbar_button_actions = {}
        
        button_names = {
            'settings': 'Settings Button',
            'refresh': 'Refresh Button',
            'telemetry': 'Telemetry Button',
            'presets': 'Presets Button',
            'compact': 'Compact Button',
            'fit': 'Fit Button',
            'connection': 'Connection Status'
        }
        
        for key, name in button_names.items():
            action = toolbar_buttons_menu.addAction(name)
            action.setCheckable(True)
            action.setChecked(self.config.toolbar_buttons_visible.get(key, True))
            action.triggered.connect(lambda checked, k=key: self._toggle_toolbar_button(k, checked))
            self.toolbar_button_actions[key] = action
        
        view_menu.addSeparator()
        view_menu.addAction("Telemetry Monitor...", self._show_telemetry)

        help_menu = menubar.addMenu("Help")
        help_menu.addAction("About", self._show_about)

    def _toggle_toolbar(self, visible: bool):
        """Toggle toolbar visibility."""
        self.config.show_toolbar = visible
        if self.toolbar:
            self.toolbar.setVisible(visible)
        self._save_config()

    def _toggle_toolbar_button(self, key: str, visible: bool):
        """Toggle individual toolbar button visibility."""
        self.config.toolbar_buttons_visible[key] = visible
        if key in self.toolbar_buttons:
            self.toolbar_buttons[key].setVisible(visible)
        self._save_config()

    def _apply_toolbar_visibility(self):
        """Apply toolbar visibility settings."""
        if self.toolbar:
            self.toolbar.setVisible(self.config.show_toolbar)
        
        for key, action in self.toolbar_buttons.items():
            visible = self.config.toolbar_buttons_visible.get(key, True)
            action.setVisible(visible)

    def _set_status(self, msg: str):
        self.statusBar().showMessage(msg)

    def _set_hint(self, msg: str):
        self.hint_label.setText(msg)

    def _show_settings(self):
        dialog = SettingsDialog(self, self.config)
        if dialog.exec() == QDialog.Accepted:
            values = dialog.get_values()
            old_outputs = self.config.num_outputs
            new_outputs = values['num_outputs']
            theme_changed = self.config.dark_theme != values['dark_theme']

            self.config.ip_address = values['ip_address']
            self.config.port = values['port']
            self.config.num_inputs = values['num_inputs']
            self.config.num_outputs = new_outputs
            self.config.label_font_family = values['label_font_family']
            self.config.label_font_size = values['label_font_size']
            self.config.button_font_family = values['button_font_family']
            self.config.button_font_size = values['button_font_size']
            self.config.active_route_color = values['active_route_color']
            self.config.dark_theme = values['dark_theme']
            self.config.crosshair_enabled = values['crosshair_enabled']
            self.config.crosshair_luminance_shift = values['crosshair_luminance_shift']
            self.config.crosshair_border_color = values['crosshair_border_color']
            self.config.use_custom_ranges = values['use_custom_ranges']
            self.config.custom_inputs = values['custom_inputs']
            self.config.custom_outputs = values['custom_outputs']
            self.config.show_input_numbers = values['show_input_numbers']
            self.config.show_output_numbers = values['show_output_numbers']

            if theme_changed:
                self._apply_theme()

            if new_outputs != old_outputs:
                self._adjust_groups_for_output_change(old_outputs, new_outputs)

            self.protocol = ETLProtocol(self.config.ip_address, self.config.port)
            self.conn_label.setText(f" {self.config.ip_address}")
            self.matrix_widget.config = self.config
            self.matrix_widget.protocol = self.protocol
            self.matrix_widget.rebuild()
            self._save_config()

    def _adjust_groups_for_output_change(self, old_outputs: int, new_outputs: int):
        if new_outputs > old_outputs:
            for out in range(old_outputs + 1, new_outputs + 1):
                self.config.output_groups.append(OutputGroup(f"Out {out}", "#b0b0b0", [out]))
        elif new_outputs < old_outputs:
            new_groups = []
            for group in self.config.output_groups:
                valid_outputs = [o for o in group.outputs if o <= new_outputs]
                if valid_outputs:
                    group.outputs = valid_outputs
                    new_groups.append(group)
            self.config.output_groups = new_groups
            if not self.config.output_groups:
                for out in range(1, new_outputs + 1):
                    self.config.output_groups.append(OutputGroup(f"Out {out}", "#b0b0b0", [out]))

    def _fit_to_screen(self):
        """Shrink window to fit within screen bounds."""
        screen = QApplication.primaryScreen().geometry()
        max_w = int(screen.width() * 0.95)
        max_h = int(screen.height() * 0.90)
        
        current_w = self.width()
        current_h = self.height()
        
        new_w = min(current_w, max_w)
        new_h = min(current_h, max_h)
        
        if self.matrix_widget:
            self.matrix_widget.setMinimumSize(0, 0)
            for btn in self.matrix_widget.route_buttons.values():
                btn.setMinimumSize(10, 15)
        
        if new_w != current_w or new_h != current_h:
            self.resize(new_w, new_h)
            self.statusBar().showMessage(f"Window resized to fit screen ({new_w}×{new_h})")
        else:
            self.statusBar().showMessage("Window already fits screen")

    def _refresh_status(self, silent=False):
        """Refresh routing status from router(s) and update the matrix display."""
        if self.config.combine_routers and self.additional_protocols:
            self._refresh_all_routers_status(silent)
        else:
            if not silent:
                self.statusBar().showMessage("Refreshing...")
            
            def do_refresh():
                try:
                    status = self.protocol.get_status() if self.protocol else None
                    routes = {}
                    if status:
                        match = re.search(r'\{BASTATUS,([^}]+)\}', status)
                        if match:
                            parts = match.group(1).split(',')
                            for i, part in enumerate(parts):
                                if part.isdigit():
                                    routes[i + 1] = int(part)
                    
                    self.refresh_complete.emit(routes, silent)
                except Exception as e:
                    print(f"Refresh error: {e}")
                    self.refresh_error.emit(str(e), silent)
            
            threading.Thread(target=do_refresh, daemon=True).start()

    def _refresh_all_routers_status(self, silent=False):
        """Refresh routing status from all routers and merge results."""
        if not silent:
            self.statusBar().showMessage("Refreshing...")
        
        def do_refresh():
            try:
                all_routes = {}
                
                # Query primary router
                status = self.protocol.get_status() if self.protocol else None
                if status:
                    match = re.search(r'\{BASTATUS,([^}]+)\}', status)
                    if match:
                        parts = match.group(1).split(',')
                        for i, part in enumerate(parts):
                            if part.isdigit():
                                all_routes[i + 1] = int(part)
                
                # Query additional routers and offset their output numbers
                if self.config.combine_routers and self.additional_protocols:
                    output_offset = self.config.primary_num_outputs or 16
                    for idx, add_protocol in enumerate(self.additional_protocols):
                        status = add_protocol.get_status()
                        if status:
                            match = re.search(r'\{BASTATUS,([^}]+)\}', status)
                            if match:
                                parts = match.group(1).split(',')
                                for i, part in enumerate(parts):
                                    if part.isdigit():
                                        all_routes[output_offset + i + 1] = int(part)
                        
                        router_outputs = self.config.additional_routers[idx].get('num_outputs', 0)
                        output_offset += router_outputs
                
                self.refresh_complete.emit(all_routes, silent)
            except Exception as e:
                print(f"Refresh error: {e}")
                self.refresh_error.emit(str(e), silent)
        
        threading.Thread(target=do_refresh, daemon=True).start()

    def _poll_all_routers(self):
        """Background poll all routers and merge status."""
        def do_poll():
            try:
                all_routes = {}
                
                # Query primary router
                if self.protocol:
                    status = self.protocol.get_status()
                    if status:
                        match = re.search(r'\{BASTATUS,([^}]+)\}', status)
                        if match:
                            parts = match.group(1).split(',')
                            for i, part in enumerate(parts):
                                if part.isdigit():
                                    all_routes[i + 1] = int(part)
                
                # Query additional routers
                if self.config.combine_routers and self.additional_protocols:
                    output_offset = self.config.primary_num_outputs or 16
                    for idx, add_protocol in enumerate(self.additional_protocols):
                        try:
                            status = add_protocol.get_status()
                            if status:
                                match = re.search(r'\{BASTATUS,([^}]+)\}', status)
                                if match:
                                    parts = match.group(1).split(',')
                                    for i, part in enumerate(parts):
                                        if part.isdigit():
                                            all_routes[output_offset + i + 1] = int(part)
                        except Exception as e:
                            print(f"Error polling router {idx + 2}: {e}")
                        
                        if idx < len(self.config.additional_routers):
                            router_outputs = self.config.additional_routers[idx].get('num_outputs', 0)
                            output_offset += router_outputs
                
                # Update UI on main thread using signal (thread-safe)
                if all_routes:
                    self.refresh_complete.emit(all_routes, True)
                    
            except Exception as e:
                print(f"Poll error: {e}")
        
        threading.Thread(target=do_poll, daemon=True).start()

    def _on_refresh_complete(self, routes: dict, silent: bool):
        """Handle refresh completion on main thread."""
        if routes:
            self._update_matrix_routes(routes)
        if not silent:
            msg = f"Status refreshed ({len(routes)} outputs)" if routes else "Could not refresh"
            self.statusBar().showMessage(msg)

    def _on_refresh_error(self, error: str, silent: bool):
        """Handle refresh error on main thread."""
        if not silent:
            self.statusBar().showMessage(f"Refresh error: {error}")

    def _update_matrix_routes(self, routes: dict):
        """Update the matrix widget with new route data."""
        if self.matrix_widget:
            self.matrix_widget.update_routes_from_telemetry(routes)

    def _show_telemetry(self):
        if self.telemetry_window is None or not self.telemetry_window.isVisible():
            self.telemetry_window = TelemetryWindow(self, self.protocol)
            self.telemetry_window.telemetry_thread.signals.status_received.connect(
                self.matrix_widget.update_routes_from_telemetry
            )
        self.telemetry_window.show()
        self.telemetry_window.raise_()

    def _toggle_compact_mode(self):
        """Toggle compact mode on/off."""
        self.config.compact_mode = self.compact_btn.isChecked()
        self.matrix_widget.rebuild()
        self._save_config()

    def _check_connection_status(self):
        """Check if router is reachable and update indicator."""
        if not self.protocol:
            return
        
        if not hasattr(self, 'conn_status_indicator'):
            return
        
        def do_check():
            try:
                response = self.protocol.get_status()
                connected = response is not None and len(response) > 0 and '{' in response
            except Exception:
                connected = False
            
            self.connection_status_changed.emit(connected)
        
        threading.Thread(target=do_check, daemon=True).start()

    def _apply_connection_indicator(self, connected: bool):
        """Apply the connection indicator update on main thread."""
        try:
            if connected:
                self.conn_status_indicator.setStyleSheet("color: #00cc00; font-size: 16px;")
                self.conn_status_indicator.setToolTip("Connected to router")
            else:
                self.conn_status_indicator.setStyleSheet("color: #cc0000; font-size: 16px;")
                self.conn_status_indicator.setToolTip("Router not reachable")
        except RuntimeError:
            pass

    def _show_presets_menu(self):
        """Show the presets dropdown menu."""
        from PySide6.QtWidgets import QMenu, QInputDialog
        
        menu = QMenu(self)
        
        save_menu = menu.addMenu("Save Preset")
        save_all = save_menu.addAction("Save All Routes...")
        save_menu.addSeparator()
        
        for group in self.config.output_groups:
            if len(group.outputs) > 1:
                save_menu.addAction(f"Save '{group.name}' Routes...").setData(('save_group', group))
        
        menu.addSeparator()
        
        if self.config.route_presets:
            load_menu = menu.addMenu("Load Preset")
            for preset in self.config.route_presets:
                scope = "All" if preset.outputs is None else f"{len(preset.outputs)} outputs"
                action = load_menu.addAction(f"{preset.name} ({scope})")
                action.setData(('load', preset))
            
            menu.addSeparator()
            
            delete_menu = menu.addMenu("Delete Preset")
            for preset in self.config.route_presets:
                action = delete_menu.addAction(preset.name)
                action.setData(('delete', preset))
        else:
            no_presets = menu.addAction("No saved presets")
            no_presets.setEnabled(False)
        
        menu.addSeparator()
        menu.addAction("Export Routes to CSV...", self._export_routes_csv)
        
        action = menu.exec_(self.cursor().pos())
        
        if action and action.data():
            cmd, data = action.data()
            if cmd == 'load':
                self._load_preset(data)
            elif cmd == 'delete':
                self._delete_preset(data)
            elif cmd == 'save_group':
                self._save_preset_for_group(data)
        elif action == save_all:
            self._save_preset_all()

    def _save_preset_all(self):
        """Save all current routes as a preset."""
        from PySide6.QtWidgets import QInputDialog
        
        if not self.matrix_widget.current_routes:
            QMessageBox.warning(self, "No Routes", "No routes to save. Please refresh status first.")
            return
        
        name, ok = QInputDialog.getText(self, "Save Preset", "Enter preset name:")
        if ok and name:
            preset = RoutePreset(
                name=name,
                routes=dict(self.matrix_widget.current_routes),
                outputs=None
            )
            self.config.route_presets.append(preset)
            self._save_config()
            self.statusBar().showMessage(f"Saved preset '{name}' with {len(preset.routes)} routes")

    def _save_preset_for_group(self, group: OutputGroup):
        """Save routes for a specific output group."""
        from PySide6.QtWidgets import QInputDialog
        
        routes = {out: inp for out, inp in self.matrix_widget.current_routes.items() 
                if out in group.outputs}
        
        if not routes:
            QMessageBox.warning(self, "No Routes", f"No routes for '{group.name}'. Please refresh status first.")
            return
        
        default_name = f"{group.name} Preset"
        name, ok = QInputDialog.getText(self, "Save Preset", "Enter preset name:", text=default_name)
        if ok and name:
            preset = RoutePreset(
                name=name,
                routes=routes,
                outputs=list(group.outputs)
            )
            self.config.route_presets.append(preset)
            self._save_config()
            self.statusBar().showMessage(f"Saved preset '{name}' with {len(routes)} routes")

    def _load_preset(self, preset: RoutePreset):
        """Load a preset and apply its routes."""
        routes_to_apply = list(preset.routes.items())
        
        if not routes_to_apply:
            return
        
        reply = QMessageBox.question(self, "Load Preset",
            f"Apply preset '{preset.name}'?\n\nThis will route {len(routes_to_apply)} outputs.",
            QMessageBox.Yes | QMessageBox.No)
        
        if reply != QMessageBox.Yes:
            return
        
        self.statusBar().showMessage(f"Applying preset '{preset.name}'...")
        
        def do_apply():
            success_count = 0
            for out, inp in routes_to_apply:
                if self.protocol.route(inp, out):
                    success_count += 1
                time.sleep(0.1)
            
            def update():
                self.statusBar().showMessage(f"✓ Applied {success_count}/{len(routes_to_apply)} routes from '{preset.name}'")
                self._refresh_status(silent=True)
            
            QTimer.singleShot(0, update)
        
        threading.Thread(target=do_apply, daemon=True).start()

    def _delete_preset(self, preset: RoutePreset):
        """Delete a preset."""
        reply = QMessageBox.question(self, "Delete Preset",
            f"Delete preset '{preset.name}'?",
            QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            self.config.route_presets = [p for p in self.config.route_presets if p.name != preset.name]
            self._save_config()
            self.statusBar().showMessage(f"Deleted preset '{preset.name}'")

    def _export_routes_csv(self):
        """Export current routes to CSV."""
        filepath, _ = QFileDialog.getSaveFileName(self, "Export Routes", "", "CSV files (*.csv)")
        if filepath:
            try:
                with open(filepath, 'w') as f:
                    f.write("Output,Input,Output Name,Input Name\n")
                    for out in sorted(self.matrix_widget.current_routes.keys()):
                        inp = self.matrix_widget.current_routes[out]
                        out_name = ""
                        for group in self.config.output_groups:
                            if out in group.outputs:
                                out_name = group.name
                                break
                        inp_name = self.config.input_names.get(inp, f"Input {inp}")
                        f.write(f"{out},{inp},{out_name},{inp_name}\n")
                self.statusBar().showMessage(f"Exported {len(self.matrix_widget.current_routes)} routes to {os.path.basename(filepath)}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not export: {e}")

    def _save_config(self):
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config.to_dict(), f, indent=2)
            self.statusBar().showMessage("Configuration saved")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not save: {e}")

    def _reset_to_defaults(self):
        """Reset all settings to defaults and restart the application."""
        reply = QMessageBox.warning(
            self,
            "Reset to Defaults",
            "This will delete all settings, groups, presets, and customizations.\n\n"
            "The application will restart with the initial setup wizard.\n\n"
            "Are you sure you want to continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            if self.bg_poll_thread:
                self.bg_poll_thread.stop()
                self.bg_poll_thread.wait(1000)
                self.bg_poll_thread = None
            
            if hasattr(self, 'conn_check_timer') and self.conn_check_timer:
                self.conn_check_timer.stop()
            
            import subprocess
            python = sys.executable
            script = os.path.abspath(sys.argv[0])
            subprocess.Popen([python, script, '--reset'])
            QApplication.quit()

    def _load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    self.config = RouterConfig.from_dict(json.load(f))
            except:
                pass

    def _load_config_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Load Configuration", "", "JSON files (*.json)")
        if filepath:
            try:
                with open(filepath, 'r') as f:
                    self.config = RouterConfig.from_dict(json.load(f))
                self.config.first_run = False
                self.protocol = ETLProtocol(self.config.ip_address, self.config.port)
                self.conn_label.setText(f" {self.config.ip_address}")
                self.matrix_widget.config = self.config
                self.matrix_widget.protocol = self.protocol
                self.matrix_widget.rebuild()
                self.statusBar().showMessage(f"Loaded {os.path.basename(filepath)}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not load: {e}")

    def _export_config(self):
        filepath, _ = QFileDialog.getSaveFileName(self, "Export Configuration", "", "JSON files (*.json)")
        if filepath:
            try:
                with open(filepath, 'w') as f:
                    json.dump(self.config.to_dict(), f, indent=2)
                self.statusBar().showMessage(f"Exported to {os.path.basename(filepath)}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not export: {e}")

    def _show_about(self):
        QMessageBox.about(self, "About",
            "ETL Vortex Matrix Controller\n\n"
            "A portable application for controlling ETL Vortex matrix routers.\n\n"
            "Tips:\n"
            "• Click output header then another to create group\n"
            "• Right-click headers to rename/color/ungroup\n"
            "• Right-click input labels to rename or adjust brightness\n"
            "• Click buttons to route, Ctrl/Shift+click to multi-select\n"
            "  - Press Enter to route selected, Escape to clear selection\n"
            "• Use Presets to save and recall routing configurations\n"
            "• Use Compact mode for large matrices\n"
            "• Green/red indicator shows router connection status")

    def _apply_theme(self):
        """Apply dark or light theme to the application."""
        app = QApplication.instance()
        palette = QPalette()
        
        if self.config.dark_theme:
            palette.setColor(QPalette.Window, QColor(53, 53, 53))
            palette.setColor(QPalette.WindowText, QColor(255, 255, 255))
            palette.setColor(QPalette.Base, QColor(35, 35, 35))
            palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
            palette.setColor(QPalette.ToolTipBase, QColor(25, 25, 25))
            palette.setColor(QPalette.ToolTipText, QColor(255, 255, 255))
            palette.setColor(QPalette.Text, QColor(255, 255, 255))
            palette.setColor(QPalette.Button, QColor(53, 53, 53))
            palette.setColor(QPalette.ButtonText, QColor(255, 255, 255))
            palette.setColor(QPalette.BrightText, QColor(255, 0, 0))
            palette.setColor(QPalette.Link, QColor(42, 130, 218))
            palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
            palette.setColor(QPalette.HighlightedText, QColor(35, 35, 35))
        else:
            palette.setColor(QPalette.Window, QColor(240, 240, 240))
            palette.setColor(QPalette.WindowText, QColor(0, 0, 0))
            palette.setColor(QPalette.Base, QColor(255, 255, 255))
            palette.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
            palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 220))
            palette.setColor(QPalette.ToolTipText, QColor(0, 0, 0))
            palette.setColor(QPalette.Text, QColor(0, 0, 0))
            palette.setColor(QPalette.Button, QColor(240, 240, 240))
            palette.setColor(QPalette.ButtonText, QColor(0, 0, 0))
            palette.setColor(QPalette.BrightText, QColor(255, 0, 0))
            palette.setColor(QPalette.Link, QColor(0, 0, 255))
            palette.setColor(QPalette.Highlight, QColor(76, 163, 224))
            palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        
        app.setPalette(palette)
        
        if hasattr(self, 'hint_label') and self.hint_label:
            hint_color = "#f0f0f0" if self.config.dark_theme else "#535353"
            self.hint_label.setStyleSheet(f"color: {hint_color}; padding: 0 10px;")

        if self.matrix_widget:
            self.matrix_widget.rebuild()

    def closeEvent(self, event):
        if hasattr(self, 'conn_check_timer') and self.conn_check_timer:
            self.conn_check_timer.stop()
        if hasattr(self, 'bg_poll_timer') and self.bg_poll_timer:
            self.bg_poll_timer.stop()
        if hasattr(self, 'bg_poll_thread') and self.bg_poll_thread:
            self.bg_poll_thread.stop()
            self.bg_poll_thread.wait(2000)
        if self.telemetry_window:
            self.telemetry_window.close()
        self._save_config()
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()