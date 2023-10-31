export interface ITestDialogContentProps {
    cancel: () => Promise<void>;
    save: (people: string[]) => Promise<void>;
}