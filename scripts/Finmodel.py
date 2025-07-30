import wb_report
import wb_prices

def main():
    """Run detailed report import followed by price loading."""
    wb_report.import_wb_detailed_reports()
    wb_prices.main()

if __name__ == "__main__":
    main()
