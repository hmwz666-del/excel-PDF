import multiprocessing
import os
import tempfile
import unittest
from queue import Empty
from unittest.mock import patch

import worker


class FakeProcess:
    def __init__(self, target=None, args=None):
        self.target = target
        self.args = args or ()
        self.started = False
        self.joined = False

    def start(self):
        self.started = True

    def is_alive(self):
        return False

    def join(self):
        self.joined = True

    def terminate(self):
        raise AssertionError("ConversionManager should not force-terminate workers")


class WorkerTests(unittest.TestCase):
    def drain_queue(self, queue_obj):
        items = []
        while True:
            try:
                items.append(queue_obj.get(timeout=0.05))
            except Empty:
                return items

    def test_worker_init_failure_leaves_shared_queue_intact(self):
        task_queue = multiprocessing.Queue()
        result_queue = multiprocessing.Queue()
        stop_event = multiprocessing.Event()

        tasks = [
            ("a.xlsx", "out", "in"),
            ("b.xlsx", "out", "in"),
            None,
            None,
        ]

        try:
            for task in tasks:
                task_queue.put(task)

            with patch("worker.ExcelConverter") as converter_cls:
                converter_cls.return_value.initialize.side_effect = RuntimeError("boom")
                worker.worker_process(task_queue, result_queue, stop_event, 1)

            self.assertEqual(self.drain_queue(task_queue), tasks)
            self.assertEqual(self.drain_queue(result_queue), [])
        finally:
            task_queue.close()
            result_queue.close()
            task_queue.join_thread()
            result_queue.join_thread()

    def test_manager_marks_pending_files_failed_when_workers_exit_early(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = os.path.join(temp_dir, "input")
            output_dir = os.path.join(temp_dir, "output")
            os.makedirs(input_dir, exist_ok=True)

            files = [
                os.path.join(input_dir, "a.xlsx"),
                os.path.join(input_dir, "b.xlsx"),
            ]

            with patch("worker.Process", FakeProcess), patch(
                "worker.scan_excel_files",
                return_value=files,
            ), patch("worker.RESULT_QUEUE_TIMEOUT", 0.01):
                manager = worker.ConversionManager(num_workers=2)
                success, failed, skipped, results = manager.start_conversion(input_dir, output_dir)

            self.assertEqual(success, 0)
            self.assertEqual(failed, 2)
            self.assertEqual(skipped, 0)
            self.assertEqual(len(results), 2)
            self.assertTrue(
                all(result.message == "工作进程提前退出，文件未被处理" for result in results)
            )


if __name__ == "__main__":
    unittest.main()
