import pandas as pd
from django.core.management.base import BaseCommand
from trials.models import ClinicalTrial
from django.utils import timezone

class Command(BaseCommand):
    help = 'Export clinical trials to Excel'

    def handle(self, *args, **options):
        trials = ClinicalTrial.objects.all().values()
        df = pd.DataFrame(trials)
        
        # Convert timezone-aware datetimes to naive
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.tz_localize(None)
        
        df.to_excel('clinical_trials.xlsx', index=False, engine='openpyxl')
        self.stdout.write(self.style.SUCCESS('Exported to clinical_trials.xlsx'))