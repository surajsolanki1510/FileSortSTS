from datetime import date

from app import (
    calculate_age_from_dob,
    clean_blood_group,
    clean_dob,
    clean_phone,
    merge_name_fields,
)


def check(actual, expected, label):
    if actual != expected:
        raise AssertionError(f"{label} failed: expected {expected!r}, got {actual!r}")


def run():
    check(clean_dob("11/06/1989"), ("11-06-1989", None), "DOB dd-mm parsing")
    check(clean_dob("5/15/1991"), ("15-05-1991", None), "DOB obvious mm-dd conversion")
    check(clean_dob("13-13-2005")[1], "red", "DOB invalid month/day detection")
    check(
        calculate_age_from_dob("11-06-1989", date(2026, 7, 7)),
        "37",
        "Age as-on calculation",
    )
    check(clean_phone("98207141969")[1], "yellow", "11-digit phone remains invalid")
    check(clean_phone("9876543210"), ("9876543210", None), "Valid phone")
    check(clean_blood_group("AB-Positive"), ("AB+", None), "Blood normalize")
    check(
        merge_name_fields("Suraj Solanki", "Suraj", "Umesh SOlanki"),
        ("Suraj", "Umesh Solanki", "Suraj Umesh Solanki"),
        "Name merge preference",
    )

    print("All regression checks passed.")


if __name__ == "__main__":
    run()
