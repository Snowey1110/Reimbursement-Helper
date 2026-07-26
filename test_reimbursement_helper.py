import tempfile
import unittest
import tempfile
from pathlib import Path

import reimbursement_helper as rh


def receipt(
    item_id: str,
    place: str,
    amount: str = "10.00",
    item_date: str = "2026-07-20",
    details: str = "",
) -> rh.ReceiptItem:
    return rh.ReceiptItem(
        item_id=item_id,
        path=f"{item_id}.png",
        filename=f"{item_id}.png",
        date=item_date,
        place=place,
        amount=amount,
        details=details,
        document_type="receipt",
        receipt_images=[{"id": item_id, "path": f"{item_id}.png", "filename": f"{item_id}.png"}],
    )


class UsaWorkflowTests(unittest.TestCase):
    def test_equal_date_and_amount_do_not_merge_without_transaction_evidence(self) -> None:
        first = receipt("first", "Vendor One", details="Parking at north garage")
        second = receipt("second", "Vendor Two", details="Parking at south garage")

        merged, _mapping, count = rh.merge_same_usa_receipt_items([first, second])

        self.assertEqual(0, count)
        self.assertEqual(["first", "second"], [item.item_id for item in merged])

    def test_matching_vendor_and_details_can_merge_same_transaction(self) -> None:
        first = receipt("first", "Embassy Suites", details="Hotel booking confirmation 441122")
        second = receipt("second", "Embassy Suites", details="Hotel booking confirmation 441122")

        merged, mapping, count = rh.merge_same_usa_receipt_items([first, second])

        self.assertEqual(1, count)
        self.assertEqual(1, len(merged))
        self.assertEqual("first", mapping["second"])

    def test_embedded_payment_proof_pages_split_as_one_proof_group(self) -> None:
        parent = receipt("parent", "Indiana Toll Road", amount="2.50")
        parent.receipt_images = [
            {
                "id": "proof-1",
                "path": "proof-page-1.png",
                "filename": "proof__RH1_t-payment_proof_d-2026-07-20_a-2p50USD__.png",
                "source_path": "proof.pdf",
                "source_page": "1",
            },
            {
                "id": "proof-2",
                "path": "proof-page-2.png",
                "filename": "proof__RH1_t-payment_proof_d-2026-07-20_a-2p50USD__.png",
                "source_path": "proof.pdf",
                "source_page": "2",
            },
            {"id": "receipt", "path": "receipt.png", "filename": "receipt.png"},
        ]

        split = rh.split_embedded_usa_payment_proofs([parent])
        receipts = [item for item in split if rh.is_usa_receipt_item(item)]
        proofs = [item for item in split if rh.is_usa_payment_proof_item(item)]

        self.assertEqual(1, len(receipts))
        self.assertEqual(1, len(rh.ensure_receipt_images(receipts[0])))
        self.assertEqual(1, len(proofs))
        self.assertEqual(2, len(rh.ensure_receipt_images(proofs[0])))
        self.assertEqual(parent.item_id, proofs[0].matched_receipt_id)

    @unittest.skipIf(rh.Image is None, "Pillow is unavailable")
    def test_duplicate_proof_is_excluded_and_one_day_match_is_allowed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            image_path = Path(directory) / "proof.png"
            rh.Image.new("RGB", (320, 180), "white").save(image_path)
            expense = receipt("expense", "Embassy Suites", amount="238.22", item_date="2026-07-09")
            proofs = []
            for item_id in ("proof-1", "proof-2"):
                proofs.append(
                    rh.ReceiptItem(
                        item_id=item_id,
                        path=str(image_path),
                        filename=f"{item_id}.png",
                        date="2026-07-08",
                        place="Chase Travel",
                        amount="238.22",
                        document_type="payment_proof",
                    )
                )
            items = [expense, *proofs]

            rh.match_usa_payment_proof_items(items)
            rows = rh.build_usa_report_rows(items)

            self.assertEqual(expense.item_id, proofs[0].matched_receipt_id)
            self.assertEqual("Duplicate proof", proofs[1].status)
            self.assertEqual(1, len(rows))
            self.assertEqual(1, len(rows[0]["proof_images"]))

    def test_manual_rotation_lock_wins_over_ai_hint(self) -> None:
        item = receipt("locked", "Vendor")
        item.receipt_images[0]["rotation_locked"] = True
        item.receipt_images[0]["rotation_degrees"] = 0
        item.rotation_locked = True

        changed = rh.apply_primary_receipt_rotation_hint(item, 90)

        self.assertFalse(changed)
        self.assertEqual(0, item.rotation_degrees)

    def test_page_results_classify_bank_portal_and_keep_pdf_page_order(self) -> None:
        item = receipt("chase", "ITR Concession Company", amount="2.50")
        item.document_type = "unknown"
        item.receipt_images = [
            {"id": "page-2", "path": "page-2.png", "filename": "page-2.png", "source_path": "proof.pdf", "source_page": "2"},
            {"id": "page-1", "path": "page-1.png", "filename": "page-1.png", "source_path": "proof.pdf", "source_page": "1"},
        ]

        changed = rh.apply_ai_page_results(
            item,
            [
                {
                    "page_index": 1,
                    "role": "payment_proof",
                    "rotation_degrees": 0,
                    "evidence": "bank_portal",
                    "transaction_date": "2026-07-20",
                    "bank_name": "Chase",
                    "card_last4": "6954",
                },
                {
                    "page_index": 2,
                    "role": "payment_proof",
                    "rotation_degrees": 0,
                    "evidence": "bank_portal",
                    "transaction_date": "2026-07-20",
                    "bank_name": "Chase",
                    "card_last4": "6954",
                },
            ],
        )

        self.assertTrue(changed)
        self.assertEqual("payment_proof", item.document_type)
        self.assertEqual(["1", "2"], [image["source_page"] for image in rh.ensure_receipt_images(item)])
        self.assertTrue(all(image["support_role"] == "payment_proof" for image in item.receipt_images))

    def test_manual_document_type_and_page_rotation_override_ai(self) -> None:
        item = receipt("paper", "Indiana Toll Road", amount="2.50")
        item.locked_document_type = "receipt"
        item.receipt_images[0].update(
            support_role="receipt",
            role_source="manual",
            rotation_degrees=270,
            rotation_locked=True,
        )

        rh.apply_ai_page_results(
            item,
            [
                {
                    "page_index": 1,
                    "role": "payment_proof",
                    "rotation_degrees": 90,
                    "evidence": "bank_portal",
                }
            ],
        )

        self.assertEqual("receipt", item.document_type)
        self.assertEqual("receipt", item.receipt_images[0]["support_role"])
        self.assertEqual(270, item.receipt_images[0]["rotation_degrees"])
        self.assertEqual("manual", item.receipt_images[0]["role_source"])

    def test_legacy_manual_memory_is_a_regeneration_hint_not_a_lock(self) -> None:
        item = receipt("legacy", "BP", amount="9.00", item_date="2026-07-06")
        item.document_type = "payment_proof"

        rh.apply_receipt_memory_record_to_item(
            item,
            {
                "version": 1,
                "reviewed_source": "manual",
                "document_type": "payment_proof",
                "date": "2026-07-06",
                "amount": "9.00",
                "page_states": [
                    {
                        "source_page": "",
                        "support_role": "payment_proof",
                        "role_source": "manual",
                        "page_amount": "9.00",
                    }
                ],
            },
            "USA",
            full_record=True,
        )

        self.assertEqual("Needs regen", item.status)
        self.assertEqual("", item.locked_document_type)
        self.assertEqual([], item.manual_fields)
        self.assertEqual("unknown", item.receipt_images[0]["support_role"])
        self.assertEqual("", item.receipt_images[0]["role_source"])

    def test_current_page_identity_replaces_stale_filename_metadata(self) -> None:
        item = receipt("fuel", "Stale vendor", amount="9.00", item_date="2026-07-06")
        item.document_type = "payment_proof"

        rh.apply_ai_page_results(
            item,
            [
                {
                    "page_index": 1,
                    "role": "receipt",
                    "rotation_degrees": 0,
                    "evidence": "photographed_paper",
                    "transaction_date": "2026-06-29",
                    "amount": "48.18",
                    "vendor": "BP",
                }
            ],
        )

        self.assertEqual("receipt", item.document_type)
        self.assertEqual("2026-06-29", item.date)
        self.assertEqual("48.18", item.amount)
        self.assertEqual("BP", item.place)
        self.assertEqual("receipt", item.receipt_images[0]["support_role"])

    def test_page_identity_does_not_overwrite_a_manually_edited_amount(self) -> None:
        item = receipt("manual-amount", "BP", amount="40.00", item_date="2026-06-29")
        item.manual_fields = ["amount"]

        rh.apply_ai_page_results(
            item,
            [
                {
                    "page_index": 1,
                    "role": "receipt",
                    "transaction_date": "2026-06-29",
                    "amount": "48.18",
                    "vendor": "BP",
                }
            ],
        )

        self.assertEqual("40.00", item.amount)

    def test_export_audit_blocks_unreviewed_and_mismatched_page_identity(self) -> None:
        unknown = receipt("unknown-page", "Vendor")
        unknown.receipt_images[0]["support_role"] = "unknown"
        legacy_known = receipt("legacy-known", "Vendor")
        legacy_known.receipt_images[0]["support_role"] = "receipt"
        mismatch = receipt("mismatch", "BP", amount="9.00", item_date="2026-07-06")
        mismatch.receipt_images[0].update(
            support_role="receipt",
            page_amount="48.18",
            transaction_date="2026-06-29",
        )

        issues = rh.usa_export_audit_issues([unknown, legacy_known, mismatch])

        self.assertIn("Pages need AI review", [reason for item, reason in issues if item is unknown])
        self.assertIn("Pages need AI review", [reason for item, reason in issues if item is legacy_known])
        self.assertIn("Page amount does not match row", [reason for item, reason in issues if item is mismatch])
        self.assertIn("Page date does not match row", [reason for item, reason in issues if item is mismatch])

        mismatch.manual_fields = ["amount", "date"]
        manual_issues = rh.usa_export_audit_issues([mismatch])
        self.assertNotIn("Page amount does not match row", [reason for _item, reason in manual_issues])
        self.assertNotIn("Page date does not match row", [reason for _item, reason in manual_issues])

    def test_companion_ticket_merges_with_same_toll_transaction(self) -> None:
        ticket = receipt("ticket", "Indiana Toll Road", amount="2.50", details="Entry Portage ticket")
        ticket.receipt_images[0]["support_role"] = "companion_ticket"
        paid = receipt("paid", "Indiana Toll Road", amount="2.50", details="Toll from Portage to LaPorte")

        merged, mapping, count = rh.merge_same_usa_receipt_items([ticket, paid])

        self.assertEqual(1, count)
        self.assertEqual(1, len(merged))
        self.assertEqual(2, len(rh.ensure_receipt_images(merged[0])))
        self.assertEqual(ticket.item_id, mapping[paid.item_id])

    def test_ambiguous_payment_proof_is_blocked_by_export_audit(self) -> None:
        first = receipt("first", "North Garage", amount="30.00")
        second = receipt("second", "South Garage", amount="30.00")
        proof = rh.ReceiptItem(
            item_id="proof",
            path="proof.png",
            filename="Chase proof.png",
            date="2026-07-20",
            amount="30.00",
            document_type="payment_proof",
            document_evidence="bank_portal",
        )
        items = [first, second, proof]

        rh.match_usa_payment_proof_items(items)
        issues = rh.usa_export_audit_issues(items)

        self.assertEqual("Multiple possible matches", proof.status)
        self.assertTrue(any(item is proof for item, _reason in issues))

    @unittest.skipIf(rh.Image is None, "Pillow is unavailable")
    def test_legacy_landscape_auto_rotation_is_reset(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            image_path = Path(directory) / "receipt__RH1_t-receipt__.png"
            rh.Image.new("RGB", (400, 200), "white").save(image_path)
            attachment = {
                "path": str(image_path),
                "filename": image_path.name,
                "rotation_degrees": 90,
            }

            self.assertTrue(rh.reset_legacy_landscape_auto_rotation(attachment))
            self.assertEqual(0, attachment["rotation_degrees"])


if __name__ == "__main__":
    unittest.main()
