import tempfile
import unittest
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
