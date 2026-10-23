// SPEC-137 "Opening a large workbook shows progress": the library's status
// line from the file choice until navigation — "Opening {title}… {n}%" as
// the workbook's row chunks arrive — and its rendering as
// spreadsheet-open-progress (role="status") in LibraryDialogs.
import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import LibraryDialogs from "./LibraryDialogs.vue";
import { spreadsheetOpenProgressText, useSpreadsheetOpenProgress } from "./useSpreadsheetOpenProgress";

const file = (name: string): File => new File([new Uint8Array(4)], name);

describe("the workbook open progress (SPEC-137)", () => {
  it("reads 'Opening {title}… {n}%', whole percent, clamped", () => {
    expect(spreadsheetOpenProgressText("Sales", 0)).toBe("Opening Sales… 0%");
    expect(spreadsheetOpenProgressText("Sales", 0.4999)).toBe("Opening Sales… 49%");
    expect(spreadsheetOpenProgressText("Sales", 1)).toBe("Opening Sales… 100%");
    expect(spreadsheetOpenProgressText("Sales", 1.5)).toBe("Opening Sales… 100%");
  });

  it("shows from the choice of a file named like a workbook package, then follows the chunks", () => {
    const progress = useSpreadsheetOpenProgress();
    expect(progress.text.value).toBe("");
    const options = progress.begin(file("Year of transactions.xlsx"));
    expect(progress.text.value).toBe("Opening Year of transactions… 0%");
    options.onSpreadsheetProgress?.(0.25);
    expect(progress.text.value).toBe("Opening Year of transactions… 25%");
    options.onSpreadsheetProgress?.(1);
    expect(progress.text.value).toBe("Opening Year of transactions… 100%");
    progress.end();
    expect(progress.text.value).toBe("");
  });

  it("names every package variant (.xlsm, .xltx, .xltm, .ods, .ots) a workbook from the choice", () => {
    for (const name of ["a.xlsm", "a.XLTX", "a.xltm", "a.ods", "a.ots"]) {
      const progress = useSpreadsheetOpenProgress();
      progress.begin(file(name));
      expect(progress.text.value, name).toBe("Opening a… 0%");
    }
  });

  it("shows for any other file only once the funnel reports a workbook (it routes by content)", () => {
    const progress = useSpreadsheetOpenProgress();
    const options = progress.begin(file("report.docx"));
    expect(progress.text.value).toBe("");
    const misnamed = useSpreadsheetOpenProgress();
    misnamed.begin(file("export.bin")).onSpreadsheetProgress?.(0);
    expect(misnamed.text.value).toBe("Opening export.bin… 0%");
    expect(options).toBeTypeOf("object");
  });

  it("never moves backwards, and a superseded open's callbacks are ignored", () => {
    const progress = useSpreadsheetOpenProgress();
    const first = progress.begin(file("one.xlsx"));
    first.onSpreadsheetProgress?.(0.5);
    first.onSpreadsheetProgress?.(0.2);
    expect(progress.text.value).toBe("Opening one… 50%");
    progress.end();
    const second = progress.begin(file("two.ods"));
    first.onSpreadsheetProgress?.(0.9);
    expect(progress.text.value).toBe("Opening two… 0%");
    second.onSpreadsheetProgress?.(0.1);
    expect(progress.text.value).toBe("Opening two… 10%");
  });

  it("renders as spreadsheet-open-progress, role=status, only while there is text", async () => {
    const wrapper = mount(LibraryDialogs, { props: { openProgress: "" } });
    expect(wrapper.find('[data-testid="spreadsheet-open-progress"]').exists()).toBe(false);
    await wrapper.setProps({ openProgress: "Opening Sales… 42%" });
    const status = wrapper.get('[data-testid="spreadsheet-open-progress"]');
    expect(status.attributes("role")).toBe("status");
    expect(status.text()).toBe("Opening Sales… 42%");
    wrapper.unmount();
  });
});
