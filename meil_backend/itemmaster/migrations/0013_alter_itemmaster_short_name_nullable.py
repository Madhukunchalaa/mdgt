from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('itemmaster', '0012_itemmaster_spec_fields'),
    ]

    operations = [
        migrations.AlterField(
            model_name='itemmaster',
            name='short_name',
            field=models.CharField(blank=True, max_length=40, null=True),
        ),
    ]
