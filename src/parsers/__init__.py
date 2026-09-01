from .apsf_activity_parser import ApsfActivityParser
from .apsf_competitor_quarterly_parser import ApsfCompetitorQuarterlyParser
from .apsf_competitor_summary_parser import ApsfCompetitorSummaryParser


FAMILY_REGISTRY = {
    "apsf_activity": ApsfActivityParser,
    "apsf_competitor_quarterly": ApsfCompetitorQuarterlyParser,
    "apsf_competitor_summary": ApsfCompetitorSummaryParser,
}

__all__ = [
    "ApsfActivityParser",
    "ApsfCompetitorQuarterlyParser",
    "ApsfCompetitorSummaryParser",
    "FAMILY_REGISTRY",
]
