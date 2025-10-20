"""
程序隔離機制
"""
import os
import tempfile
import shutil
from typing import Dict, Any, Optional, Callable, List, Tuple
from pathlib import Path
from dataclasses import dataclass
import threading
import time
from contextlib import contextmanager

from ..exceptions import RenderError


@dataclass
class IsolationConfig:
    """隔離配置"""
    use_temp_directory: bool = True
    cleanup_on_error: bool = True
    max_memory_mb: Optional[int] = None
    timeout_seconds: Optional[float] = None
    backup_original: bool = True


@dataclass
class ProcessState:
    """程序狀態"""
    process_id: str
    start_time: float
    temp_dir: Optional[str]
    original_file: Optional[str]
    backup_file: Optional[str]
    is_completed: bool = False
    has_error: bool = False
    error_message: Optional[str] = None


class ProcessIsolationManager:
    """
    程序隔離管理器
    
    提供安全的檔案處理環境，確保渲染過程不會影響原始檔案
    """
    
    def __init__(self, config: Optional[IsolationConfig] = None):
        """
        初始化隔離管理器
        
        Args:
            config: 隔離配置，如果為None則使用預設配置
        """
        self.config = config or IsolationConfig()
        self._active_processes: Dict[str, ProcessState] = {}
        self._lock = threading.Lock()
        self._cleanup_callbacks: Dict[str, List[Callable]] = {}
    
    @contextmanager
    def isolated_process(self, template_path: str, process_id: Optional[str] = None):
        """
        建立隔離程序環境的上下文管理器
        
        Args:
            template_path: 模板檔案路徑
            process_id: 程序ID，如果為None則自動生成
            
        Yields:
            Tuple[str, str]: (工作檔案路徑, 程序ID)
        """
        if process_id is None:
            process_id = self._generate_process_id()
        
        process_state = None
        try:
            # 設置隔離環境
            working_file, process_state = self._setup_isolation(template_path, process_id)
            
            yield working_file, process_id
            
            # 標記完成
            if process_state:
                process_state.is_completed = True
                
        except Exception as e:
            # 處理錯誤
            if process_state:
                process_state.has_error = True
                process_state.error_message = str(e)
            
            if self.config.cleanup_on_error:
                self._cleanup_process(process_id)
            
            raise
        
        finally:
            # 清理資源
            if process_state and process_state.is_completed and not process_state.has_error:
                self._cleanup_process(process_id)
    
    def _setup_isolation(self, template_path: str, process_id: str) -> Tuple[str, ProcessState]:
        """
        設置隔離環境
        
        Args:
            template_path: 模板檔案路徑
            process_id: 程序ID
            
        Returns:
            Tuple[str, ProcessState]: (工作檔案路徑, 程序狀態)
        """
        with self._lock:
            if process_id in self._active_processes:
                raise RenderError(f"程序 {process_id} 已經在執行中")
            
            # 建立程序狀態
            process_state = ProcessState(
                process_id=process_id,
                start_time=time.time(),
                temp_dir=None,
                original_file=template_path,
                backup_file=None
            )
            
            try:
                # 建立臨時目錄
                if self.config.use_temp_directory:
                    temp_dir = tempfile.mkdtemp(prefix=f"excel_render_{process_id}_")
                    process_state.temp_dir = temp_dir
                    
                    # 複製檔案到臨時目錄
                    template_name = Path(template_path).name
                    working_file = os.path.join(temp_dir, template_name)
                    shutil.copy2(template_path, working_file)
                else:
                    working_file = template_path
                
                # 建立備份
                if self.config.backup_original:
                    backup_file = self._create_backup(template_path, process_id)
                    process_state.backup_file = backup_file
                
                # 註冊程序
                self._active_processes[process_id] = process_state
                self._cleanup_callbacks[process_id] = []
                
                return working_file, process_state
                
            except Exception as e:
                # 清理已建立的資源
                if process_state.temp_dir and os.path.exists(process_state.temp_dir):
                    shutil.rmtree(process_state.temp_dir, ignore_errors=True)
                raise RenderError(f"設置隔離環境失敗: {str(e)}")
    
    def _create_backup(self, file_path: str, process_id: str) -> str:
        """
        建立檔案備份
        
        Args:
            file_path: 原始檔案路徑
            process_id: 程序ID
            
        Returns:
            str: 備份檔案路徑
        """
        file_path_obj = Path(file_path)
        timestamp = int(time.time())
        backup_name = f"{file_path_obj.stem}_{process_id}_{timestamp}{file_path_obj.suffix}"
        backup_path = file_path_obj.parent / f"backup_{backup_name}"
        
        shutil.copy2(file_path, backup_path)
        return str(backup_path)
    
    def _cleanup_process(self, process_id: str) -> None:
        """
        清理程序資源
        
        Args:
            process_id: 程序ID
        """
        with self._lock:
            if process_id not in self._active_processes:
                return
            
            process_state = self._active_processes[process_id]
            
            try:
                # 執行清理回調
                if process_id in self._cleanup_callbacks:
                    for callback in self._cleanup_callbacks[process_id]:
                        try:
                            callback()
                        except Exception:
                            # 忽略清理回調中的錯誤
                            pass
                
                # 清理臨時目錄
                if process_state.temp_dir and os.path.exists(process_state.temp_dir):
                    shutil.rmtree(process_state.temp_dir, ignore_errors=True)
                
                # 清理備份檔案（如果程序成功完成）
                if (process_state.backup_file and 
                    os.path.exists(process_state.backup_file) and 
                    process_state.is_completed and 
                    not process_state.has_error):
                    os.remove(process_state.backup_file)
                
            except Exception:
                # 忽略清理過程中的錯誤
                pass
            
            finally:
                # 移除程序記錄
                del self._active_processes[process_id]
                if process_id in self._cleanup_callbacks:
                    del self._cleanup_callbacks[process_id]
    
    def register_cleanup_callback(self, process_id: str, callback: Callable) -> None:
        """
        註冊清理回調函數
        
        Args:
            process_id: 程序ID
            callback: 清理回調函數
        """
        with self._lock:
            if process_id in self._cleanup_callbacks:
                self._cleanup_callbacks[process_id].append(callback)
    
    def get_process_state(self, process_id: str) -> Optional[ProcessState]:
        """
        取得程序狀態
        
        Args:
            process_id: 程序ID
            
        Returns:
            Optional[ProcessState]: 程序狀態，如果不存在則返回None
        """
        with self._lock:
            return self._active_processes.get(process_id)
    
    def get_active_processes(self) -> Dict[str, ProcessState]:
        """
        取得所有活躍程序
        
        Returns:
            Dict[str, ProcessState]: 程序ID到狀態的映射
        """
        with self._lock:
            return self._active_processes.copy()
    
    def force_cleanup_process(self, process_id: str) -> bool:
        """
        強制清理程序
        
        Args:
            process_id: 程序ID
            
        Returns:
            bool: 是否成功清理
        """
        try:
            self._cleanup_process(process_id)
            return True
        except Exception:
            return False
    
    def cleanup_all_processes(self) -> int:
        """
        清理所有程序
        
        Returns:
            int: 清理的程序數量
        """
        with self._lock:
            process_ids = list(self._active_processes.keys())
        
        cleaned_count = 0
        for process_id in process_ids:
            if self.force_cleanup_process(process_id):
                cleaned_count += 1
        
        return cleaned_count
    
    def _generate_process_id(self) -> str:
        """
        生成程序ID
        
        Returns:
            str: 程序ID
        """
        import uuid
        return f"proc_{int(time.time())}_{uuid.uuid4().hex[:8]}"
    
    def check_memory_usage(self, process_id: str) -> Optional[float]:
        """
        檢查程序記憶體使用量
        
        Args:
            process_id: 程序ID
            
        Returns:
            Optional[float]: 記憶體使用量（MB），如果無法檢查則返回None
        """
        try:
            # 嘗試導入psutil，如果沒有安裝則忽略
            import psutil
            current_process = psutil.Process()
            memory_info = current_process.memory_info()
            return memory_info.rss / (1024 * 1024)  # 轉換為MB
        except ImportError:
            # psutil未安裝，無法檢查記憶體
            return None
        except Exception:
            # 其他錯誤
            return None
    
    def check_timeout(self, process_id: str) -> bool:
        """
        檢查程序是否超時
        
        Args:
            process_id: 程序ID
            
        Returns:
            bool: 是否超時
        """
        if self.config.timeout_seconds is None:
            return False
        
        process_state = self.get_process_state(process_id)
        if not process_state:
            return False
        
        elapsed_time = time.time() - process_state.start_time
        return elapsed_time > self.config.timeout_seconds
    
    def restore_from_backup(self, process_id: str) -> bool:
        """
        從備份還原檔案
        
        Args:
            process_id: 程序ID
            
        Returns:
            bool: 是否成功還原
        """
        process_state = self.get_process_state(process_id)
        if not process_state or not process_state.backup_file:
            return False
        
        try:
            if (os.path.exists(process_state.backup_file) and 
                process_state.original_file):
                shutil.copy2(process_state.backup_file, process_state.original_file)
                return True
        except Exception:
            pass
        
        return False
    
    def get_isolation_stats(self) -> Dict[str, Any]:
        """
        取得隔離統計資訊
        
        Returns:
            Dict[str, Any]: 統計資訊
        """
        with self._lock:
            active_count = len(self._active_processes)
            completed_count = sum(1 for p in self._active_processes.values() if p.is_completed)
            error_count = sum(1 for p in self._active_processes.values() if p.has_error)
            
            return {
                "active_processes": active_count,
                "completed_processes": completed_count,
                "error_processes": error_count,
                "config": {
                    "use_temp_directory": self.config.use_temp_directory,
                    "cleanup_on_error": self.config.cleanup_on_error,
                    "backup_original": self.config.backup_original,
                    "max_memory_mb": self.config.max_memory_mb,
                    "timeout_seconds": self.config.timeout_seconds
                }
            }
