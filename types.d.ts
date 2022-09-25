
type Judge = 'vjudge' | 'codeforces' | 'atcoder'

interface VjudgeContest {
  id: string;
  title?: string;
  reqCount: number;
  reqProblems: string[]
}

interface User {
  id: string;
  name: string;
  handles: Record<Judge, string[]>
}

type UserSolvedMap = Record<number, number[]>

interface VjudgeResponseData {
  id: number;
  title: string;
  begin: number;
  length: number;
  isReplay: boolean;
  version: string;
  participants: Record<string, string[]>
  submissions: number[][];
}

type color = 'green' | 'yellow' | 'orange' | 'red';
