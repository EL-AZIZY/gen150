from .apsf_activity_parser import ApsfActivityParser
from .apsf_competitor_summary_parser import ApsfCompetitorSummaryParser
from .apsf_competitor_quarterly_parser import ApsfCompetitorQuarterlyParser
from .flat_table_parser import FlatTableParser
from .production_marque_penetration_parser import ProductionMarquePenetrationParser
from .production_marque_concession_parser import ProductionMarqueConcessionParser
from .production_marque_pivot_staging_parser import ProductionMarquePivotStagingParser
from .production_marque_variation_tpr_parser import ProductionMarqueVariationTprParser


FAMILY_REGISTRY = {
    "apsf_activity": ApsfActivityParser,
    "apsf_competitor_summary": ApsfCompetitorSummaryParser,
    "apsf_competitor_quarterly": ApsfCompetitorQuarterlyParser,
    "production_marque_details": FlatTableParser,
    "production_marque_penetration": ProductionMarquePenetrationParser,
    "production_marque_concession": ProductionMarqueConcessionParser,
    "production_marque_pivot_staging": ProductionMarquePivotStagingParser,
    "production_marque_variation_tpr": ProductionMarqueVariationTprParser,
}

__all__ = [
    "ApsfActivityParser",
    "ApsfCompetitorSummaryParser",
    "ApsfCompetitorQuarterlyParser",
    "FlatTableParser",
    "ProductionMarquePenetrationParser",
    "ProductionMarqueConcessionParser",
    "ProductionMarquePivotStagingParser",
    "ProductionMarqueVariationTprParser",
    "FAMILY_REGISTRY",
]
