import { describe, it, expect } from "vitest";
import { findMatchingPatient, findEvidence, evaluateEvidence } from "./reviewEngine";

describe("Patient Matching", () => {
    it("matches a patient using name and DOB", () => {
        const payerRow = {
            name: "John Doe",
            dob: "01/02/1960"
        };

        const emr = [
            {
            "Patient Name": "John Doe",
            "Date of Birth": "01/02/1960"
            }
        ];

        const mapping = {
            name: "Patient Name",
            dob: "Date of Birth"
        };

        const result = findMatchingPatient(
            payerRow,
            emr,
            mapping,
            "name",
            "dob"
        );

        expect(result).toEqual(emr[0]);
    })

    it("does not match patients with names", () => {
        const payerRow = {
            name: "John Smith",
            dob: "01/02/1060"
        };

        const emr = [
            {
            "Patient Name": "John Doe",
            "Date of Birth": "01/02/1060"
            }
        ];

        const mapping = {
            name: "Patient Name",
            dob: "Date of Birth"
        };

        const result = findMatchingPatient(
            payerRow,
            emr,
            mapping,
            "name",
            "dob"
        );

        expect(result).toBeUndefined();
    })

    it("does not match patients with dobs", () => {
        const payerRow = {
            name: "John Doe",
            dob: "01/02/1960"
        };

        const emr = [
            {
            "Patient Name": "John Doe",
            "Date of Birth": "01/02/1959"
            }
        ];

        const mapping = {
            name: "Patient Name",
            dob: "Date of Birth"
        };

        const result = findMatchingPatient(
            payerRow,
            emr,
            mapping,
            "name",
            "dob"
        );

        expect(result).toBeUndefined();
    })

    it("does not match a patient when the EMR DOB is missing", () => {
        const payerRow = {
            name: "John Doe",
            dob: "08/03/1982"
        };

        const emr = [
            {
                "Patient Name": "John Doe",
                "Date of Birth": "" 
            }
        ]

        const mapping = {
            name: "Patient Name",
            dob: "Date of Birth"
        };

        const result = findMatchingPatient(
            payerRow,
            emr,
            mapping,
            "name",
            "dob"
        )

        expect(result).toBeUndefined();
    });
});

describe("Clinical Evidence", () => {

});