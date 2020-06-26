# Mesurements must be given in meters as it is the reference unit.
# Ratio =  1 inches / 0.0254 Meters

import sys
from api import Api, raw_input


def correct_bom():
    try:
        session = Api()
        print("- Author: recs@premiertech.com")
        print("- Last update: 2020-05-01")
        session.check_valid_version("Solid Edge ST7", "Solid Edge 2019")
        draft = session.active_document()
        print("Document Number: %s\n" % draft.name)
        assert draft.name.lower().endswith(".dft"), (
            "This macro only works on draft document not %s" % draft.name[-4:]
        )

        if draft.PartsLists.Count == 1:

            print("\nMax number of rows:")
            print(
                "  \t- First Page..................: %s"
                % draft.PartsLists[1].MaximumRowsFirstPage
            )
            print(
                "  \t- Additional Page.............: %s"
                % draft.PartsLists[1].MaximumRowsAdditionalPages
            )

            print("\nMax height:")
            mhfp = draft.PartsLists[1].MaximumHeightFirstPage
            mhap = draft.PartsLists[1].MaximumHeightAdditionalPages
            print("\t- First Page..................: %s" % mhfp)
            print("\t- Additional Page.............: %s" % mhap)
            print("\n")
            print("\t\t****************")

            response = raw_input(
                "Would you like to overwrite the rows with 100-100? [y/Y]es:\n>"
            )
            if response.lower() in ["y", "yes"]:
                pass
            else:
                sys.exit()
            # =========== Overwrite ===============

            draft.PartsLists[1].MaximumRowsFirstPage = 100
            draft.PartsLists[1].MaximumRowsAdditionalPages = 100
            print("Overwrite rows with 100-100: \tdone")

            mhfp, mhap = 0.1778, 0.1778
            draft.PartsLists[1].MaximumHeightFirstPage = mhfp
            draft.PartsLists[1].MaximumHeightAdditionalPages = mhap
            print("Overwrite heights: \tdone")

            print("\n")
            print("\t\t****************")
            print("\n")
            print("Max number of rows:")
            print(
                "\t- First Page..................: %s"
                % draft.PartsLists[1].MaximumRowsFirstPage
            )
            print(
                "\t- Additional Page.............: %s"
                % draft.PartsLists[1].MaximumRowsAdditionalPages
            )
            print("\n")
            print("Max height:")
            print(
                "\t- First Page..................: %s"
                % draft.PartsLists[1].MaximumHeightFirstPage
            )
            print(
                "\t- Additional Page.............: %s"
                % draft.PartsLists[1].MaximumHeightAdditionalPages
            )

        else:
            pass

    except AssertionError as err:
        print(err.args)
    except Exception as ex:
        print(ex.args)
    finally:
        raw_input("\nPress any key to exit...")
        sys.exit()


def confirmation(func):
    response = raw_input("""Correct the display of BOM for manuals, (Press y/[Y] to proceed.): """)
    if response.lower() not in ["y"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    confirmation(correct_bom)
