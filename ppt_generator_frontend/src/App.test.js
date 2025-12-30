import { render, screen } from "@testing-library/react";
import App from "./App";

test("renders ppt generator header", () => {
  render(<App />);
  const heading = screen.getByText(/PowerPoint Generator/i);
  expect(heading).toBeInTheDocument();
});

test("renders generate button", () => {
  render(<App />);
  const button = screen.getByRole("button", { name: /Generate PPT/i });
  expect(button).toBeInTheDocument();
});
