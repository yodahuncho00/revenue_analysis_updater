import os
import json
from datetime import datetime
from typing import Dict, Any

class RunReportGenerator:
    """Generates structured JSON and human-readable Markdown run reports."""

    @staticmethod
    def generate_report(
        run_id: str,
        success: bool,
        date_range: Dict[str, Any],
        row_counts: Dict[str, int],
        checks_passed: list,
        checks_failed: list,
        duration_seconds: float,
        error_message: str = None,
        output_dir: str = "data/reports"
    ) -> Dict[str, Any]:
        os.makedirs(output_dir, exist_ok=True)
        report_data = {
            "run_id": run_id,
            "timestamp": datetime.now().isoformat(),
            "status": "SUCCESS" if success else "FAILED",
            "duration_seconds": round(duration_seconds, 2),
            "date_range": date_range,
            "row_counts": row_counts,
            "checks_passed": checks_passed,
            "checks_failed": checks_failed,
            "error_message": error_message
        }

        # Write JSON report
        json_path = os.path.join(output_dir, f"{run_id}.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)

        # Write Markdown report
        md_path = os.path.join(output_dir, f"{run_id}.md")
        status_badge = "✅ SUCCESS" if success else "❌ FAILED"
        md_content = f"""# Daily Revenue Analysis ETL Run Report

- **Run ID:** `{run_id}`
- **Status:** {status_badge}
- **Timestamp:** {report_data['timestamp']}
- **Duration:** {report_data['duration_seconds']}s
- **Date Range:** From `{date_range.get('from_date')}` to `{date_range.get('to_date')}` (Reporting Month: `{date_range.get('rpt_month_end')}`)

## Row Counts by Sheet
| Sheet | Rows Written / Appended |
|---|---|
| `BO_extract` | {row_counts.get('bo_extract', 0)} |
| `Data` | {row_counts.get('data', 0)} |
| `BO_merchant_balance` | {row_counts.get('bo_merchant_balance', 0)} |
| `BO_provider_balance` | {row_counts.get('bo_provider_balance', 0)} |
| `BO_settlements` | {row_counts.get('bo_settlements', 0)} |
| `BO_topups` | {row_counts.get('bo_topups', 0)} |

## Integrity Verification Gates
- **Checks Passed ({len(checks_passed)}):**
"""
        for chk in checks_passed:
            md_content += f"  - ✅ {chk}\n"

        if checks_failed:
            md_content += f"\n- **Checks Failed ({len(checks_failed)}):**\n"
            for chk in checks_failed:
                md_content += f"  - ❌ {chk}\n"

        if error_message:
            md_content += f"\n## Failure Diagnostic\n```\n{error_message}\n```\n"

        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)

        return report_data
