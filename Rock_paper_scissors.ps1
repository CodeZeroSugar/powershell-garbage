function Play-RockPaperScissors {

$player = Read-Host "Choose rock, paper, or scissors"
$player = $player.ToLower()

$choices = "rock","paper","scissors"

if($choices -notcontains $player){
    do{
        $input = Read-Host "Choose rock, paper, or scissors"
        $player = $input
        $player = $player.ToLower()
        }
    while($choices -notcontains $player)
}

$computer = Get-Random -InputObject "rock","paper","scissors"
Write-host "Computer chose $computer..." -ForegroundColor Red



if($computer -eq $player)
    {Write-Host "Tie!" -ForegroundColor Yellow}

if($computer -eq "rock"){
    if($player -eq "paper"){
        Write-Host "You win!" -ForegroundColor Cyan
        }
    elseif($player -eq "scissors"){
        Write-Host "You lose!" -ForegroundColor Red
    }
 }

if($computer -eq "paper"){
    if($player -eq "scissors"){
        Write-Host "You win!" -ForegroundColor Cyan
        }
    elseif($player -eq "rock"){
        Write-Host "You lose!" -ForegroundColor Red
    }
 }

 if($computer -eq "scissors"){
    if($player -eq "rock"){
        Write-Host "You win!" -ForegroundColor Cyan
        }
    elseif($player -eq "paper"){
        Write-Host "You lose!" -ForegroundColor Red
    }
  }
}

 

do{    
    Play-RockPaperScissors
    $response = Read-Host "Would you like to play again? (Yes/No)"
    }
until($response -eq "no")

