from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('itemmaster', '0010_alter_itemmaster_sap_item_id'),
    ]

    operations = [
        migrations.AlterField(
            model_name='itemmaster',
            name='short_name',
            field=models.CharField(max_length=40),
        ),
    ]
