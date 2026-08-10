// SPEC-041 AC-7, AC-9, AC-11 — custom property name/value validation.
import { describe, expect, it } from "vitest";
import {
  canonicalCustomPropertyValue,
  createCustomProperty,
  customPropertyNameError,
  customPropertyValueError,
  defaultCustomPropertyValue,
} from "./customProperties";

describe("customPropertyNameError (AC-7)", () => {
  it("rejects an empty name", () => {
    expect(customPropertyNameError("", [])).not.toBe("");
    expect(customPropertyNameError("   ", [])).not.toBe("");
  });

  it("rejects a duplicate name", () => {
    expect(customPropertyNameError("Client", ["Client", "Budget"])).not.toBe("");
  });

  it("accepts a unique, non-empty name", () => {
    expect(customPropertyNameError("Client", ["Budget"])).toBe("");
  });
});

describe("customPropertyValueError (AC-9, AC-11)", () => {
  it("text accepts anything, including empty", () => {
    expect(customPropertyValueError("text", "")).toBe("");
    expect(customPropertyValueError("text", "anything at all")).toBe("");
  });

  it("number rejects non-numeric text (AC-11's own example)", () => {
    expect(customPropertyValueError("number", "not a number")).not.toBe("");
    expect(customPropertyValueError("number", "")).not.toBe("");
  });

  it("number accepts integers and decimals", () => {
    expect(customPropertyValueError("number", "42")).toBe("");
    expect(customPropertyValueError("number", "-3.5")).toBe("");
  });

  it("date rejects malformed and out-of-range dates", () => {
    expect(customPropertyValueError("date", "not a date")).not.toBe("");
    expect(customPropertyValueError("date", "2024-02-30")).not.toBe(""); // no such day
    expect(customPropertyValueError("date", "2024/02/01")).not.toBe(""); // wrong separator
  });

  it("date accepts a real YYYY-MM-DD date", () => {
    expect(customPropertyValueError("date", "2024-02-29")).toBe(""); // leap year
  });

  it("boolean only accepts true/false", () => {
    expect(customPropertyValueError("boolean", "true")).toBe("");
    expect(customPropertyValueError("boolean", "false")).toBe("");
    expect(customPropertyValueError("boolean", "yes")).not.toBe("");
  });
});

describe("canonicalCustomPropertyValue", () => {
  it("normalises a number's string form", () => {
    expect(canonicalCustomPropertyValue("number", " 42.50 ")).toBe("42.5");
  });

  it("trims text/date/boolean", () => {
    expect(canonicalCustomPropertyValue("text", "  hi  ")).toBe("hi");
    expect(canonicalCustomPropertyValue("date", " 2024-01-01 ")).toBe("2024-01-01");
  });
});

describe("defaultCustomPropertyValue", () => {
  it("defaults boolean to false and everything else to empty", () => {
    expect(defaultCustomPropertyValue("boolean")).toBe("false");
    expect(defaultCustomPropertyValue("text")).toBe("");
    expect(defaultCustomPropertyValue("number")).toBe("");
    expect(defaultCustomPropertyValue("date")).toBe("");
  });
});

describe("createCustomProperty", () => {
  it("trims the name and canonicalises the value", () => {
    expect(createCustomProperty("  Client ", "number", " 42.50 ")).toEqual({
      name: "Client",
      type: "number",
      value: "42.5",
    });
  });
});
