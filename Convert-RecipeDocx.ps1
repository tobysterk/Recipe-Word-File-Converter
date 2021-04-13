[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string] $wordDocFolder
)

function Write-YamlFile
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $RecipeObject
    )
    

    $RecipeObject.Properties.GetEnumer

    $sb = New-Object -TypeName System.Text.StringBuilder
    $sb.AppendLine("name: $($RecipeObject.Title)")
    $sb.AppendLine("servings: $($RecipeObject.Servings) servings")
    $sb.AppendLine("source: $($RecipeObject.Source)")
    $sb.AppendLine("on_favorites: no")
    $sb.AppendLine("prep_time: 0 minutes")
    $sb.AppendLine("cook_time: 0 minutes")
<#     
    notes: |
      7/28: no nuts, 50 min-ish bake in 9x5 loaf pan
    ingredients: |
      2 c flour
      3 c grated carrots (1 lb)
      2 c sugar (for this I reduced it to 3/4 cup white and 3/4 cup brown)
      1.25 c oil
      2 tsp cinnamon (I added more than 2 tsp. though cause I prefer a strong cinnamon flavour)
      1 tsp salt
      1 tsp baking soda
      4 eggs
      1 c nuts (I used roasted walnuts)
    directions: |
      Peel then finely grate the carrots. Since carrots will release a lot of moisture when baked, squeeze some of the juice out.
      Mix the dry ingredients (flour, baking soda, salt, cinnamon, and nuts). Aside from cinnamon, I also added a dash of nutmeg.
      In a separate bowl, mix the wet ingredients (oil and sugar). Beat in the eggs, one at a time, into the wet mixture.*
      *So I didn't exactly follow this step after doing it for the first time. When I followed it, I had a film of oil on top of the wet mixture. Not sure how that affects the cake since it tasted fine anyway but after the first attempt, I followed Claire from Bon Appetit's method wherein she beat the eggs and sugars until pale and ribbony then slowly streamed in oil. This no longer led to me having a film of oil on top.
      Mix in the carrots. Gradually mix in the dry mixture into the wet.
      Pour onto a greased and floured pan then bake at 350 degrees for 30 mins. or until a toothpick comes out clean.
     #>
}

function Import-WordFile
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $WordFile
    )
    
    $recipeObject = New-Object -TypeName PSObject -Property
    @{
        Title = $null
        Servings = $null
        Source = $null
        $onFavorites = "no"
        $prepTime = "0 minutes"
        $cookTime = "0 minutes"
        $notes = $null
        $ingredients = $null
        $directions = $null
    }

    return $recipeObject
}

# main
if (-not (Test-Path -Path $wordDocFolder))
{
    throw "ERROR: folder path not valid"
}
$word = New-Object -ComObject Word.application

$wordFiles = Get-ChildItem -Path $wordDocFolder | Where-Object -Property Name -Match ".+\.docx"

# Write-Host $wordFiles

foreach ($file in $wordFiles)
{
    $document = $word.Documents.Open($file.FullName)
    Import-WordFile -WordFile $document
}
