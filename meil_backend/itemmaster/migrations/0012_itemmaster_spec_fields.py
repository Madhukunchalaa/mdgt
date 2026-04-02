from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('itemmaster', '0010_alter_itemmaster_sap_item_id'),
    ]

    operations = [
        migrations.AddField(
            model_name='itemmaster',
            name='item_type',
            field=models.CharField(blank=True, help_text='Type of material', max_length=100, null=True),
        ),
        migrations.AddField(
            model_name='itemmaster',
            name='item_number',
            field=models.CharField(blank=True, help_text='Item number', max_length=100, null=True),
        ),
        migrations.AddField(
            model_name='itemmaster',
            name='moc',
            field=models.CharField(blank=True, help_text='Material of Construction', max_length=100, null=True),
        ),
        migrations.AddField(
            model_name='itemmaster',
            name='item_size',
            field=models.CharField(blank=True, help_text='Size', max_length=100, null=True),
        ),
        migrations.AddField(
            model_name='itemmaster',
            name='part_number',
            field=models.CharField(blank=True, help_text='Part number', max_length=100, null=True),
        ),
        migrations.AddField(
            model_name='itemmaster',
            name='model',
            field=models.CharField(blank=True, help_text='Model', max_length=100, null=True),
        ),
        migrations.AddField(
            model_name='itemmaster',
            name='make',
            field=models.CharField(blank=True, help_text='Make / Manufacturer', max_length=100, null=True),
        ),
    ]
